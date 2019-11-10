VERSION 5.00

Begin VB.Form frmGenerateurVille
    Caption = "Générateur de ville"
    ScaleMode = 1
    AutoRedraw = 0              'False
    FontTransparent = -1              'True
    LinkTopic = "Form1"
    MDIChild = -1              'True
    ClientLeft   = 60
    ClientTop    = 450
    ClientWidth  = 6825
    ClientHeight = 8895
    BeginProperty Font
        Name          = "Times New Roman"
        Size          = 9
        Charset       = 0
        Weight        = 400
        Underline     = 0              'False
        Italic        = 0              'False
        Strikethrough = 0              'False
    EndProperty
    Begin VB.CommandButton cmbPrint
        Caption = "Imprimer"
        Left   = 4440
        Top    = 1200
        Width  = 2295
        Height = 375
        TabIndex = 12
    End
    Begin VB.CommandButton cmdGenerer
        Caption = "Générer"
        Left   = 120
        Top    = 1200
        Width  = 2295
        Height = 375
        TabIndex = 11
    End
    Begin VB.TextBox fldstrSortie
        Left   = 120
        Top    = 1680
        Width  = 6650
        Height = 7100
        TabIndex = 10
        MultiLine = -1              'True
        ScrollBars = 2
    End
    Begin VB.TextBox fldnDime
        Left   = 6240
        Top    = 600
        Width  = 495
        Height = 375
        Text = "0"
        TabIndex = 9
    End
    Begin VB.TextBox fldnImpot
        Left   = 4800
        Top    = 600
        Width  = 495
        Height = 375
        Text = "0"
        TabIndex = 8
    End
    Begin VB.TextBox fldstrnom
        Left   = 1920
        Top    = 120
        Width  = 1455
        Height = 375
        Text = "Nom"
        TabIndex = 5
    End
    Begin VB.ComboBox CmbType
        Style = 2
        Left   = 1920
        Top    = 615
        Width  = 1935
        Height = 345
        Text = ""
        TabIndex = 3
    End
    Begin VB.TextBox fldnPopulation
        Left   = 5760
        Top    = 120
        Width  = 975
        Height = 375
        Text = "0"
        TabIndex = 0
    End
    Begin VB.Label lbl
        Caption = "% de dîme"
        Index = 4
        Left   = 5400
        Top    = 675
        Width  = 750
        Height = 225
        TabIndex = 7
        AutoSize = -1              'True
    End
    Begin VB.Label lbl
        Caption = "% d'impôt"
        Index = 3
        Left   = 3960
        Top    = 675
        Width  = 750
        Height = 225
        TabIndex = 6
        AutoSize = -1              'True
    End
    Begin VB.Label lbl
        Caption = "Nom de la communauté"
        Index = 0
        Left   = 120
        Top    = 195
        Width  = 1695
        Height = 225
        TabIndex = 4
        AutoSize = -1              'True
    End
    Begin VB.Label lbl
        Caption = "Type de communauté"
        Index = 2
        Left   = 120
        Top    = 675
        Width  = 1575
        Height = 225
        TabIndex = 2
        AutoSize = -1              'True
    End
    Begin VB.Label lbl
        Caption = "Population de la communauté"
        Index = 1
        Left   = 3480
        Top    = 195
        Width  = 2130
        Height = 225
        TabIndex = 1
        AutoSize = -1              'True
    End
End
Public Function TypeVille(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'00721c00    55                      push ebp
'00721c01    8bec                    mov ebp, esp
'00721c03    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00721c06    6866724000              push 00407266
'00721c0b    64a100000000            mov ax, word ptr fs:[00000000]
'00721c11    50                      push eax
    'Reference to '__except_list'
'00721c12    64892500000000          mov dword ptr fs:[00000000], esp
'00721c19    83ec0c                  sub esp, 0c
'00721c1c    53                      push ebx
'00721c1d    56                      push esi
'00721c1e    57                      push edi
'00721c1f    8965f4                  mov dword ptr [ebp-0c], esp
'00721c22    c745f8d86f4000          mov dword ptr [ebp-08], 00406fd8
'00721c29    33f6                    xor esi, esi
'00721c2b    8975fc                  mov dword ptr [ebp-04], esi
'00721c2e    8b4508                  mov eax, dword ptr [ebp+08]
'00721c31    8b08                    mov ecx, dword ptr [eax]
'00721c33    50                      push eax
'00721c34    ff5104                  call dword ptr [ecx+04]
'00721c37    8b5510                  mov edx, dword ptr [ebp+10]
'00721c3a    8b450c                  mov eax, dword ptr [ebp+0c]
'00721c3d    8932                    mov dword ptr [edx], esi
'00721c3f    8b00                    mov eax, dword ptr [eax]
'00721c41    83f850                  cmp eax, 50
'00721c44    8975e8                  mov dword ptr [ebp-18], esi
'00721c47    7f07                    jg 721c50

If (arg_0 <= 80) Then
'00721c49    ba146a4500              mov edx, 00456a14
'00721c4e    eb57                    jmp 721ca7
    
End If
'00721c50    3d90010000              cmp eax, 00000190
'00721c55    7f07                    jg 721c5e

If (arg_0 <= 400) Then
'00721c57    baf06b4500              mov edx, 00456bf0
'00721c5c    eb49                    jmp 721ca7
    
End If
'00721c5e    3d84030000              cmp eax, 00000384
'00721c63    7f07                    jg 721c6c

If (arg_0 <= 900) Then
'00721c65    ba10684500              mov edx, 00456810
'00721c6a    eb3b                    jmp 721ca7
    
End If
'00721c6c    3dd0070000              cmp eax, 000007d0
'00721c71    7f07                    jg 721c7a

If (arg_0 <= 2000) Then
'00721c73    ba24684500              mov edx, 00456824
'00721c78    eb2d                    jmp 721ca7
    
End If
'00721c7a    3d88130000              cmp eax, 00001388
'00721c7f    7f07                    jg 721c88

If (arg_0 <= 5000) Then
'00721c81    ba34684500              mov edx, 00456834
'00721c86    eb1f                    jmp 721ca7
    
End If
'00721c88    3de02e0000              cmp eax, 00002ee0
'00721c8d    7f07                    jg 721c96

If (arg_0 <= 12000) Then
'00721c8f    ba70664500              mov edx, 00456670
'00721c94    eb11                    jmp 721ca7
    
End If
'00721c96    3da8610000              cmp eax, 000061a8
'00721c9b    ba90664500              mov edx, 00456690
'00721ca0    7e05                    jle 721ca7

If (arg_0 > 25000) Then
'00721ca2    baa0644500              mov edx, 004564a0
    
End If
'00721ca7    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaStrCopy"
'00721caa    ff1548124000            call dword ptr [00401248]
var_41 = ("Cité")
'00721cb0    68c21c7200              push 00721cc2
'00721cb5    eb0a                    jmp 721cc1
'00721cb7    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'00721cba    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'00721cc0    c3                      ret
'00721cc1    c3                      ret
'00721cc2    8b4508                  mov eax, dword ptr [ebp+08]
'00721cc5    8b08                    mov ecx, dword ptr [eax]
'00721cc7    50                      push eax
'00721cc8    ff5108                  call dword ptr [ecx+08]
'00721ccb    8b45e8                  mov eax, dword ptr [ebp-18]
'00721cce    8b5510                  mov edx, dword ptr [ebp+10]
'00721cd1    8902                    mov dword ptr [edx], eax
'00721cd3    8b45fc                  mov eax, dword ptr [ebp-04]
'00721cd6    8b4dec                  mov ecx, dword ptr [ebp-14]
'00721cd9    5f                      pop edi
'00721cda    5e                      pop esi
    'Reference to '__except_list'
'00721cdb    64890d00000000          mov dword ptr fs:[00000000], ecx
'00721ce2    5b                      pop ebx
'00721ce3    8be5                    mov esp, ebp
'00721ce5    5d                      pop ebp
'00721ce6    c20c00                  ret 000c


End Function


Public Function LimiteFinanciere(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'00721cf0    55                      push ebp
'00721cf1    8bec                    mov ebp, esp
'00721cf3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00721cf6    6866724000              push 00407266
'00721cfb    64a100000000            mov ax, word ptr fs:[00000000]
'00721d01    50                      push eax
    'Reference to '__except_list'
'00721d02    64892500000000          mov dword ptr fs:[00000000], esp
'00721d09    83ec0c                  sub esp, 0c
'00721d0c    53                      push ebx
'00721d0d    56                      push esi
'00721d0e    57                      push edi
'00721d0f    8965f4                  mov dword ptr [ebp-0c], esp
'00721d12    c745f8e86f4000          mov dword ptr [ebp-08], 00406fe8
'00721d19    33f6                    xor esi, esi
'00721d1b    8975fc                  mov dword ptr [ebp-04], esi
'00721d1e    8b4508                  mov eax, dword ptr [ebp+08]
'00721d21    8b08                    mov ecx, dword ptr [eax]
'00721d23    50                      push eax
'00721d24    ff5104                  call dword ptr [ecx+04]
'00721d27    8b550c                  mov edx, dword ptr [ebp+0c]
'00721d2a    8b02                    mov eax, dword ptr [edx]
'00721d2c    3bc6                    cmp eax, esi
'00721d2e    8975e8                  mov dword ptr [ebp-18], esi
'00721d31    7f05                    jg 721d38

If (arg_0 <= 0) Then
'00721d33    8975e8                  mov dword ptr [ebp-18], esi
'00721d36    eb78                    jmp 721db0
    
End If
'00721d38    83f850                  cmp eax, 50
'00721d3b    7f09                    jg 721d46

If (arg_0 <= 80) Then
'00721d3d    c745e828000000          mov dword ptr [ebp-18], 00000028
'00721d44    eb6a                    jmp 721db0
    
End If
'00721d46    3d90010000              cmp eax, 00000190
'00721d4b    7f09                    jg 721d56

If (arg_0 <= 400) Then
'00721d4d    c745e864000000          mov dword ptr [ebp-18], 00000064
'00721d54    eb5a                    jmp 721db0
    
End If
'00721d56    3d84030000              cmp eax, 00000384
'00721d5b    7f09                    jg 721d66

If (arg_0 <= 900) Then
'00721d5d    c745e8c8000000          mov dword ptr [ebp-18], 000000c8
'00721d64    eb4a                    jmp 721db0
    
End If
'00721d66    3dd0070000              cmp eax, 000007d0
'00721d6b    7f09                    jg 721d76

If (arg_0 <= 2000) Then
'00721d6d    c745e820030000          mov dword ptr [ebp-18], 00000320
'00721d74    eb3a                    jmp 721db0
    
End If
'00721d76    3d88130000              cmp eax, 00001388
'00721d7b    7f09                    jg 721d86

If (arg_0 <= 5000) Then
'00721d7d    c745e8b80b0000          mov dword ptr [ebp-18], 00000bb8
'00721d84    eb2a                    jmp 721db0
    
End If
'00721d86    3de02e0000              cmp eax, 00002ee0
'00721d8b    7f09                    jg 721d96

If (arg_0 <= 12000) Then
'00721d8d    c745e8983a0000          mov dword ptr [ebp-18], 00003a98
'00721d94    eb1a                    jmp 721db0
    
End If
'00721d96    33c9                    xor ecx, ecx
var_num3 = Empty
'00721d98    3da8610000              cmp eax, 000061a8
'00721d9d    0f9fc1                  setg cl
'00721da0    49                      dec ecx
'00721da1    81e1a015ffff            and ecx, ffff15a0
var_num3 = arg_0 Not > 25000 And -60000
'00721da7    81c1a0860100            add ecx, 000186a0
var_num3 = var_num3 + 100000
'00721dad    894de8                  mov dword ptr [ebp-18], ecx
'00721db0    8b4508                  mov eax, dword ptr [ebp+08]
'00721db3    8b10                    mov edx, dword ptr [eax]
'00721db5    50                      push eax
'00721db6    ff5208                  call dword ptr [edx+08]
'00721db9    8b4510                  mov eax, dword ptr [ebp+10]
'00721dbc    8b4de8                  mov ecx, dword ptr [ebp-18]
'00721dbf    8908                    mov dword ptr [eax], ecx
'00721dc1    8b45fc                  mov eax, dword ptr [ebp-04]
'00721dc4    8b4dec                  mov ecx, dword ptr [ebp-14]
'00721dc7    5f                      pop edi
'00721dc8    5e                      pop esi
    'Reference to '__except_list'
'00721dc9    64890d00000000          mov dword ptr fs:[00000000], ecx
'00721dd0    5b                      pop ebx
'00721dd1    8be5                    mov esp, ebp
'00721dd3    5d                      pop ebp
'00721dd4    c20c00                  ret 000c


End Function


Public Function LiquiditeDisponible(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'00721de0    55                      push ebp
'00721de1    8bec                    mov ebp, esp
'00721de3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00721de6    6866724000              push 00407266
'00721deb    64a100000000            mov ax, word ptr fs:[00000000]
'00721df1    50                      push eax
    'Reference to '__except_list'
'00721df2    64892500000000          mov dword ptr fs:[00000000], esp
'00721df9    83ec0c                  sub esp, 0c
'00721dfc    53                      push ebx
'00721dfd    56                      push esi
'00721dfe    57                      push edi
'00721dff    8965f4                  mov dword ptr [ebp-0c], esp
'00721e02    c745f8f06f4000          mov dword ptr [ebp-08], 00406ff0
'00721e09    33f6                    xor esi, esi
'00721e0b    8975fc                  mov dword ptr [ebp-04], esi
'00721e0e    8b4508                  mov eax, dword ptr [ebp+08]
'00721e11    8b08                    mov ecx, dword ptr [eax]
'00721e13    50                      push eax
'00721e14    ff5104                  call dword ptr [ecx+04]
'00721e17    8b550c                  mov edx, dword ptr [ebp+0c]
'00721e1a    8b02                    mov eax, dword ptr [edx]
'00721e1c    83f850                  cmp eax, 50
'00721e1f    8975e8                  mov dword ptr [ebp-18], esi
'00721e22    7f05                    jg 721e29

If (arg_0 <= 80) Then
'00721e24    6bc002                  imul eax, 02
    var_num1 = arg_0 * 2
'00721e27    eb57                    jmp 721e80
    
End If
'00721e29    3d90010000              cmp eax, 00000190
'00721e2e    7f05                    jg 721e35

If (arg_0 <= 400) Then
'00721e30    6bc005                  imul eax, 05
    var_num1 = arg_0 * 5
'00721e33    eb4b                    jmp 721e80
    
End If
'00721e35    3d84030000              cmp eax, 00000384
'00721e3a    7f05                    jg 721e41

If (arg_0 <= 900) Then
'00721e3c    6bc00a                  imul eax, 0a
    var_num1 = arg_0 * 10
'00721e3f    eb3f                    jmp 721e80
    
End If
'00721e41    3dd0070000              cmp eax, 000007d0
'00721e46    7f05                    jg 721e4d

If (arg_0 <= 2000) Then
'00721e48    6bc028                  imul eax, 28
    var_num1 = arg_0 * 40
'00721e4b    eb33                    jmp 721e80
    
End If
'00721e4d    3d88130000              cmp eax, 00001388
'00721e52    7f08                    jg 721e5c

If (arg_0 <= 5000) Then
'00721e54    69c096000000            imul eax, 00000096
    var_num1 = arg_0 * 150
'00721e5a    eb24                    jmp 721e80
    
End If
'00721e5c    3de02e0000              cmp eax, 00002ee0
'00721e61    7f08                    jg 721e6b

If (arg_0 <= 12000) Then
'00721e63    69c0ee020000            imul eax, 000002ee
    var_num1 = arg_0 * 750
'00721e69    eb15                    jmp 721e80
    
End If
'00721e6b    3da8610000              cmp eax, 000061a8
'00721e70    7f08                    jg 721e7a

If (arg_0 <= 25000) Then
'00721e72    69c0d0070000            imul eax, 000007d0
    var_num1 = arg_0 * 2000
'00721e78    eb06                    jmp 721e80
    
End If
'00721e7a    69c088130000            imul eax, 00001388
var_num1 = arg_0 * 5000
'00721e80    702a                    jo 721eac
'00721e82    8945e8                  mov dword ptr [ebp-18], eax
'00721e85    8b4508                  mov eax, dword ptr [ebp+08]
'00721e88    8b08                    mov ecx, dword ptr [eax]
'00721e8a    50                      push eax
'00721e8b    ff5108                  call dword ptr [ecx+08]
'00721e8e    8b45e8                  mov eax, dword ptr [ebp-18]
'00721e91    8b5510                  mov edx, dword ptr [ebp+10]
'00721e94    8902                    mov dword ptr [edx], eax
'00721e96    8b45fc                  mov eax, dword ptr [ebp-04]
'00721e99    8b4dec                  mov ecx, dword ptr [ebp-14]
'00721e9c    5f                      pop edi
'00721e9d    5e                      pop esi
    'Reference to '__except_list'
'00721e9e    64890d00000000          mov dword ptr fs:[00000000], ecx
'00721ea5    5b                      pop ebx
'00721ea6    8be5                    mov esp, ebp
'00721ea8    5d                      pop ebp
'00721ea9    c20c00                  ret 000c


End Function


Public Function Revenu(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'00721ec0    55                      push ebp
'00721ec1    8bec                    mov ebp, esp
'00721ec3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00721ec6    6866724000              push 00407266
'00721ecb    64a100000000            mov ax, word ptr fs:[00000000]
'00721ed1    50                      push eax
    'Reference to '__except_list'
'00721ed2    64892500000000          mov dword ptr fs:[00000000], esp
'00721ed9    83ec40                  sub esp, 40
'00721edc    53                      push ebx
'00721edd    56                      push esi
'00721ede    57                      push edi
'00721edf    8965f4                  mov dword ptr [ebp-0c], esp
'00721ee2    c745f828704000          mov dword ptr [ebp-08], 00407028
'00721ee9    33f6                    xor esi, esi
'00721eeb    8975fc                  mov dword ptr [ebp-04], esi
'00721eee    8b4508                  mov eax, dword ptr [ebp+08]
'00721ef1    8b08                    mov ecx, dword ptr [eax]
'00721ef3    50                      push eax
'00721ef4    ff5104                  call dword ptr [ecx+04]
'00721ef7    8b550c                  mov edx, dword ptr [ebp+0c]
'00721efa    8b02                    mov eax, dword ptr [edx]
'00721efc    83f850                  cmp eax, 50
'00721eff    8975e8                  mov dword ptr [ebp-18], esi
'00721f02    8945dc                  mov dword ptr [ebp-24], eax
'00721f05    7f24                    jg 721f2b

If (arg_0 <= 80) Then
'00721f07    db45dc                  fild dword ptr [ebp-24]
'00721f0a    dd5dd4                  fstp qword ptr [ebp-2c]
    'var_44 = (00)
'00721f0d    dd45d4                  fld qword ptr [ebp-2c]
'00721f10    dc0d20704000            fmul qword ptr [00407020]
'00721f16    dfe0                    fnstsw ax
'00721f18    a80d                    test al, 0d
'00721f1a    0f8508010000            jne 722028

' *** Reference to "__vbaFpI4"
'00721f20    ff15a4124000            call dword ptr [004012a4]
    'var_num1 = CLng((((arg_0)) * 2.5#))
'00721f26    e9d3000000              jmp 721ffe
    
End If
'00721f2b    3d90010000              cmp eax, 00000190
'00721f30    7f08                    jg 721f3a

If (arg_0 <= 400) Then
'00721f32    6bc003                  imul eax, 03
    var_num1 = arg_0 * 3
'00721f35    e9c2000000              jmp 721ffc
    
End If
'00721f3a    3d84030000              cmp eax, 00000384
'00721f3f    7f24                    jg 721f65

If (arg_0 <= 900) Then
'00721f41    db45dc                  fild dword ptr [ebp-24]
'00721f44    dd5dcc                  fstp qword ptr [ebp-34]
    'var_43 = (00)
'00721f47    dd45cc                  fld qword ptr [ebp-34]
'00721f4a    dc0d18704000            fmul qword ptr [00407018]
'00721f50    dfe0                    fnstsw ax
'00721f52    a80d                    test al, 0d
'00721f54    0f85ce000000            jne 722028

' *** Reference to "__vbaFpI4"
'00721f5a    ff15a4124000            call dword ptr [004012a4]
    'var_num1 = CLng((((arg_0)) * 3.7#))
'00721f60    e999000000              jmp 721ffe
    
End If
'00721f65    3dd0070000              cmp eax, 000007d0
'00721f6a    7f21                    jg 721f8d

If (arg_0 <= 2000) Then
'00721f6c    db45dc                  fild dword ptr [ebp-24]
'00721f6f    dd5dc4                  fstp qword ptr [ebp-3c]
    'var_9 = (00)
'00721f72    dd45c4                  fld qword ptr [ebp-3c]
'00721f75    dc0d10704000            fmul qword ptr [00407010]
'00721f7b    dfe0                    fnstsw ax
'00721f7d    a80d                    test al, 0d
'00721f7f    0f85a3000000            jne 722028

' *** Reference to "__vbaFpI4"
'00721f85    ff15a4124000            call dword ptr [004012a4]
    'var_num1 = CLng((((arg_0)) * 4.5#))
'00721f8b    eb71                    jmp 721ffe
    
End If
'00721f8d    3d88130000              cmp eax, 00001388
'00721f92    7f1d                    jg 721fb1

If (arg_0 <= 5000) Then
'00721f94    db45dc                  fild dword ptr [ebp-24]
'00721f97    dd5dbc                  fstp qword ptr [ebp-44]
    'var_58 = (00)
'00721f9a    dd45bc                  fld qword ptr [ebp-44]
'00721f9d    dc0d08704000            fmul qword ptr [00407008]
'00721fa3    dfe0                    fnstsw ax
'00721fa5    a80d                    test al, 0d
'00721fa7    757f                    jne 722028

' *** Reference to "__vbaFpI4"
'00721fa9    ff15a4124000            call dword ptr [004012a4]
    'var_num1 = CLng((((arg_0)) * 5.5#))
'00721faf    eb4d                    jmp 721ffe
    
End If
'00721fb1    3de02e0000              cmp eax, 00002ee0
'00721fb6    7f1d                    jg 721fd5

If (arg_0 <= 12000) Then
'00721fb8    db45dc                  fild dword ptr [ebp-24]
'00721fbb    dd5db4                  fstp qword ptr [ebp-4c]
    'var_62 = (00)
'00721fbe    dd45b4                  fld qword ptr [ebp-4c]
'00721fc1    dc0d00704000            fmul qword ptr [00407000]
'00721fc7    dfe0                    fnstsw ax
'00721fc9    a80d                    test al, 0d
'00721fcb    755b                    jne 722028

' *** Reference to "__vbaFpI4"
'00721fcd    ff15a4124000            call dword ptr [004012a4]
    'var_num1 = CLng((((arg_0)) * 6.7#))
'00721fd3    eb29                    jmp 721ffe
    
End If
'00721fd5    3da8610000              cmp eax, 000061a8
'00721fda    7f1d                    jg 721ff9

If (arg_0 <= 25000) Then
'00721fdc    db45dc                  fild dword ptr [ebp-24]
'00721fdf    dd5dac                  fstp qword ptr [ebp-54]
    'var_50 = (00)
'00721fe2    dd45ac                  fld qword ptr [ebp-54]
'00721fe5    dc0df86f4000            fmul qword ptr [00406ff8]
'00721feb    dfe0                    fnstsw ax
'00721fed    a80d                    test al, 0d
'00721fef    7537                    jne 722028

' *** Reference to "__vbaFpI4"
'00721ff1    ff15a4124000            call dword ptr [004012a4]
    'var_num1 = CLng((((arg_0)) * 8.2#))
'00721ff7    eb05                    jmp 721ffe
    
End If
'00721ff9    6bc00a                  imul eax, 0a
var_num1 = arg_0 * 10
'00721ffc    702f                    jo 72202d
'00721ffe    8945e8                  mov dword ptr [ebp-18], eax
'00722001    8b4508                  mov eax, dword ptr [ebp+08]
'00722004    8b08                    mov ecx, dword ptr [eax]
'00722006    50                      push eax
'00722007    ff5108                  call dword ptr [ecx+08]
'0072200a    8b45e8                  mov eax, dword ptr [ebp-18]
'0072200d    8b5510                  mov edx, dword ptr [ebp+10]
'00722010    8902                    mov dword ptr [edx], eax
'00722012    8b45fc                  mov eax, dword ptr [ebp-04]
'00722015    8b4dec                  mov ecx, dword ptr [ebp-14]
'00722018    5f                      pop edi
'00722019    5e                      pop esi
    'Reference to '__except_list'
'0072201a    64890d00000000          mov dword ptr fs:[00000000], ecx
'00722021    5b                      pop ebx
'00722022    8be5                    mov esp, ebp
'00722024    5d                      pop ebp
'00722025    c20c00                  ret 000c


End Function


Public Function RevenuOr(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'00722040    55                      push ebp
'00722041    8bec                    mov ebp, esp
'00722043    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00722046    6866724000              push 00407266
'0072204b    64a100000000            mov ax, word ptr fs:[00000000]
'00722051    50                      push eax
    'Reference to '__except_list'
'00722052    64892500000000          mov dword ptr fs:[00000000], esp
'00722059    83ec58                  sub esp, 58
'0072205c    53                      push ebx
'0072205d    56                      push esi
'0072205e    57                      push edi
'0072205f    8965f4                  mov dword ptr [ebp-0c], esp
'00722062    c745f860704000          mov dword ptr [ebp-08], 00407060
'00722069    33f6                    xor esi, esi
'0072206b    8975fc                  mov dword ptr [ebp-04], esi
'0072206e    8b4508                  mov eax, dword ptr [ebp+08]
'00722071    8b08                    mov ecx, dword ptr [eax]
'00722073    50                      push eax
'00722074    ff5104                  call dword ptr [ecx+04]
'00722077    8b550c                  mov edx, dword ptr [ebp+0c]
'0072207a    8b02                    mov eax, dword ptr [edx]
'0072207c    83f850                  cmp eax, 50
'0072207f    8975e8                  mov dword ptr [ebp-18], esi
'00722082    8945dc                  mov dword ptr [ebp-24], eax
'00722085    7f1a                    jg 7220a1

If (arg_0 <= 80) Then
'00722087    db45dc                  fild dword ptr [ebp-24]
'0072208a    dd5dd4                  fstp qword ptr [ebp-2c]
    'var_44 = (00)
'0072208d    dd45d4                  fld qword ptr [ebp-2c]
'00722090    dc0d20704000            fmul qword ptr [00407020]
'00722096    dc0d58704000            fmul qword ptr [00407058]
'0072209c    e9d7000000              jmp 722178
    
End If
'007220a1    3d90010000              cmp eax, 00000190
'007220a6    7f20                    jg 7220c8

If (arg_0 <= 400) Then
'007220a8    6bc003                  imul eax, 03
    var_num1 = arg_0 * 3
'007220ab    0f8002010000            jo 7221b3
'007220b1    8945d0                  mov dword ptr [ebp-30], eax
'007220b4    db45d0                  fild dword ptr [ebp-30]
'007220b7    dd5dc8                  fstp qword ptr [ebp-38]
    'var_46 = (00)
'007220ba    dd45c8                  fld qword ptr [ebp-38]
'007220bd    dc0d50704000            fmul qword ptr [00407050]
'007220c3    e9b0000000              jmp 722178
    
End If
'007220c8    3d84030000              cmp eax, 00000384
'007220cd    7f1a                    jg 7220e9

If (arg_0 <= 900) Then
'007220cf    db45dc                  fild dword ptr [ebp-24]
'007220d2    dd5dc0                  fstp qword ptr [ebp-40]
    'var_5 = (00)
'007220d5    dd45c0                  fld qword ptr [ebp-40]
'007220d8    dc0d18704000            fmul qword ptr [00407018]
'007220de    dc0d48704000            fmul qword ptr [00407048]
'007220e4    e98f000000              jmp 722178
    
End If
'007220e9    3dd0070000              cmp eax, 000007d0
'007220ee    7f17                    jg 722107

If (arg_0 <= 2000) Then
'007220f0    db45dc                  fild dword ptr [ebp-24]
'007220f3    dd5db8                  fstp qword ptr [ebp-48]
    'var_61 = (00)
'007220f6    dd45b8                  fld qword ptr [ebp-48]
'007220f9    dc0d10704000            fmul qword ptr [00407010]
'007220ff    dc0d40704000            fmul qword ptr [00407040]
'00722105    eb71                    jmp 722178
    
End If
'00722107    3d88130000              cmp eax, 00001388
'0072210c    7f17                    jg 722125

If (arg_0 <= 5000) Then
'0072210e    db45dc                  fild dword ptr [ebp-24]
'00722111    dd5db0                  fstp qword ptr [ebp-50]
    'var_6 = (00)
'00722114    dd45b0                  fld qword ptr [ebp-50]
'00722117    dc0d08704000            fmul qword ptr [00407008]
'0072211d    dc0d986c4000            fmul qword ptr [00406c98]
'00722123    eb53                    jmp 722178
    
End If
'00722125    3de02e0000              cmp eax, 00002ee0
'0072212a    7f17                    jg 722143

If (arg_0 <= 12000) Then
'0072212c    db45dc                  fild dword ptr [ebp-24]
'0072212f    dd5da8                  fstp qword ptr [ebp-58]
    'var_86 = (00)
'00722132    dd45a8                  fld qword ptr [ebp-58]
'00722135    dc0d00704000            fmul qword ptr [00407000]
'0072213b    dc0d20164000            fmul qword ptr [00401620]
'00722141    eb35                    jmp 722178
    
End If
'00722143    3da8610000              cmp eax, 000061a8
'00722148    7f17                    jg 722161

If (arg_0 <= 25000) Then
'0072214a    db45dc                  fild dword ptr [ebp-24]
'0072214d    dd5da0                  fstp qword ptr [ebp-60]
    'var_7 = (00)
'00722150    dd45a0                  fld qword ptr [ebp-60]
'00722153    dc0df86f4000            fmul qword ptr [00406ff8]
'00722159    dc0d38704000            fmul qword ptr [00407038]
'0072215f    eb17                    jmp 722178
    
End If
'00722161    6bc00a                  imul eax, 0a
var_num1 = arg_0 * 10
'00722164    704d                    jo 7221b3
'00722166    89459c                  mov dword ptr [ebp-64], eax
'00722169    db459c                  fild dword ptr [ebp-64]
'0072216c    dd5d94                  fstp qword ptr [ebp-6c]
'var_121 = (00)
'0072216f    dd4594                  fld qword ptr [ebp-6c]
'00722172    dc0d30704000            fmul qword ptr [00407030]
'00722178    dfe0                    fnstsw ax
'0072217a    a80d                    test al, 0d
'0072217c    7530                    jne 7221ae

' *** Reference to "__vbaFpI4"
'0072217e    ff15a4124000            call dword ptr [004012a4]
'var_num1 = CLng((((var_num1)) * 0.8#))
'00722184    8945e8                  mov dword ptr [ebp-18], eax
'00722187    8b4508                  mov eax, dword ptr [ebp+08]
'0072218a    8b08                    mov ecx, dword ptr [eax]
'0072218c    50                      push eax
'0072218d    ff5108                  call dword ptr [ecx+08]
'00722190    8b45e8                  mov eax, dword ptr [ebp-18]
'00722193    8b5510                  mov edx, dword ptr [ebp+10]
'00722196    8902                    mov dword ptr [edx], eax
'00722198    8b45fc                  mov eax, dword ptr [ebp-04]
'0072219b    8b4dec                  mov ecx, dword ptr [ebp-14]
'0072219e    5f                      pop edi
'0072219f    5e                      pop esi
    'Reference to '__except_list'
'007221a0    64890d00000000          mov dword ptr fs:[00000000], ecx
'007221a7    5b                      pop ebx
'007221a8    8be5                    mov esp, ebp
'007221aa    5d                      pop ebp
'007221ab    c20c00                  ret 000c


End Function


Public Function RevenuPremiere(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'007221c0    55                      push ebp
'007221c1    8bec                    mov ebp, esp
'007221c3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'007221c6    6866724000              push 00407266
'007221cb    64a100000000            mov ax, word ptr fs:[00000000]
'007221d1    50                      push eax
    'Reference to '__except_list'
'007221d2    64892500000000          mov dword ptr fs:[00000000], esp
'007221d9    83ec58                  sub esp, 58
'007221dc    53                      push ebx
'007221dd    56                      push esi
'007221de    57                      push edi
'007221df    8965f4                  mov dword ptr [ebp-0c], esp
'007221e2    c745f888704000          mov dword ptr [ebp-08], 00407088
'007221e9    33f6                    xor esi, esi
'007221eb    8975fc                  mov dword ptr [ebp-04], esi
'007221ee    8b4508                  mov eax, dword ptr [ebp+08]
'007221f1    8b08                    mov ecx, dword ptr [eax]
'007221f3    50                      push eax
'007221f4    ff5104                  call dword ptr [ecx+04]
'007221f7    8b550c                  mov edx, dword ptr [ebp+0c]
'007221fa    8b02                    mov eax, dword ptr [edx]
'007221fc    83f850                  cmp eax, 50
'007221ff    8975e8                  mov dword ptr [ebp-18], esi
'00722202    8945dc                  mov dword ptr [ebp-24], eax
'00722205    7f1a                    jg 722221

If (arg_0 <= 80) Then
'00722207    db45dc                  fild dword ptr [ebp-24]
'0072220a    dd5dd4                  fstp qword ptr [ebp-2c]
    'var_44 = (00)
'0072220d    dd45d4                  fld qword ptr [ebp-2c]
'00722210    dc0d20704000            fmul qword ptr [00407020]
'00722216    dc0d80704000            fmul qword ptr [00407080]
'0072221c    e9d7000000              jmp 7222f8
    
End If
'00722221    3d90010000              cmp eax, 00000190
'00722226    7f20                    jg 722248

If (arg_0 <= 400) Then
'00722228    6bc003                  imul eax, 03
    var_num1 = arg_0 * 3
'0072222b    0f8002010000            jo 722333
'00722231    8945d0                  mov dword ptr [ebp-30], eax
'00722234    db45d0                  fild dword ptr [ebp-30]
'00722237    dd5dc8                  fstp qword ptr [ebp-38]
    'var_46 = (00)
'0072223a    dd45c8                  fld qword ptr [ebp-38]
'0072223d    dc0d78704000            fmul qword ptr [00407078]
'00722243    e9b0000000              jmp 7222f8
    
End If
'00722248    3d84030000              cmp eax, 00000384
'0072224d    7f1a                    jg 722269

If (arg_0 <= 900) Then
'0072224f    db45dc                  fild dword ptr [ebp-24]
'00722252    dd5dc0                  fstp qword ptr [ebp-40]
    'var_5 = (00)
'00722255    dd45c0                  fld qword ptr [ebp-40]
'00722258    dc0d18704000            fmul qword ptr [00407018]
'0072225e    dc0d30704000            fmul qword ptr [00407030]
'00722264    e98f000000              jmp 7222f8
    
End If
'00722269    3dd0070000              cmp eax, 000007d0
'0072226e    7f17                    jg 722287

If (arg_0 <= 2000) Then
'00722270    db45dc                  fild dword ptr [ebp-24]
'00722273    dd5db8                  fstp qword ptr [ebp-48]
    'var_61 = (00)
'00722276    dd45b8                  fld qword ptr [ebp-48]
'00722279    dc0d10704000            fmul qword ptr [00407010]
'0072227f    dc0d70704000            fmul qword ptr [00407070]
'00722285    eb71                    jmp 7222f8
    
End If
'00722287    3d88130000              cmp eax, 00001388
'0072228c    7f17                    jg 7222a5

If (arg_0 <= 5000) Then
'0072228e    db45dc                  fild dword ptr [ebp-24]
'00722291    dd5db0                  fstp qword ptr [ebp-50]
    'var_6 = (00)
'00722294    dd45b0                  fld qword ptr [ebp-50]
'00722297    dc0d08704000            fmul qword ptr [00407008]
'0072229d    dc0d20674000            fmul qword ptr [00406720]
'007222a3    eb53                    jmp 7222f8
    
End If
'007222a5    3de02e0000              cmp eax, 00002ee0
'007222aa    7f17                    jg 7222c3

If (arg_0 <= 12000) Then
'007222ac    db45dc                  fild dword ptr [ebp-24]
'007222af    dd5da8                  fstp qword ptr [ebp-58]
    'var_86 = (00)
'007222b2    dd45a8                  fld qword ptr [ebp-58]
'007222b5    dc0d00704000            fmul qword ptr [00407000]
'007222bb    dc0d20164000            fmul qword ptr [00401620]
'007222c1    eb35                    jmp 7222f8
    
End If
'007222c3    3da8610000              cmp eax, 000061a8
'007222c8    7f17                    jg 7222e1

If (arg_0 <= 25000) Then
'007222ca    db45dc                  fild dword ptr [ebp-24]
'007222cd    dd5da0                  fstp qword ptr [ebp-60]
    'var_7 = (00)
'007222d0    dd45a0                  fld qword ptr [ebp-60]
'007222d3    dc0df86f4000            fmul qword ptr [00406ff8]
'007222d9    dc0d68704000            fmul qword ptr [00407068]
'007222df    eb17                    jmp 7222f8
    
End If
'007222e1    6bc00a                  imul eax, 0a
var_num1 = arg_0 * 10
'007222e4    704d                    jo 722333
'007222e6    89459c                  mov dword ptr [ebp-64], eax
'007222e9    db459c                  fild dword ptr [ebp-64]
'007222ec    dd5d94                  fstp qword ptr [ebp-6c]
'var_121 = (00)
'007222ef    dd4594                  fld qword ptr [ebp-6c]
'007222f2    dc0d48704000            fmul qword ptr [00407048]
'007222f8    dfe0                    fnstsw ax
'007222fa    a80d                    test al, 0d
'007222fc    7530                    jne 72232e

' *** Reference to "__vbaFpI4"
'007222fe    ff15a4124000            call dword ptr [004012a4]
'var_num1 = CLng((((var_num1)) * 0.2#))
'00722304    8945e8                  mov dword ptr [ebp-18], eax
'00722307    8b4508                  mov eax, dword ptr [ebp+08]
'0072230a    8b08                    mov ecx, dword ptr [eax]
'0072230c    50                      push eax
'0072230d    ff5108                  call dword ptr [ecx+08]
'00722310    8b45e8                  mov eax, dword ptr [ebp-18]
'00722313    8b5510                  mov edx, dword ptr [ebp+10]
'00722316    8902                    mov dword ptr [edx], eax
'00722318    8b45fc                  mov eax, dword ptr [ebp-04]
'0072231b    8b4dec                  mov ecx, dword ptr [ebp-14]
'0072231e    5f                      pop edi
'0072231f    5e                      pop esi
    'Reference to '__except_list'
'00722320    64890d00000000          mov dword ptr fs:[00000000], ecx
'00722327    5b                      pop ebx
'00722328    8be5                    mov esp, ebp
'0072232a    5d                      pop ebp
'0072232b    c20c00                  ret 000c


End Function


Public Function Instances(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'00722340    55                      push ebp
'00722341    8bec                    mov ebp, esp
'00722343    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00722346    6866724000              push 00407266
'0072234b    64a100000000            mov ax, word ptr fs:[00000000]
'00722351    50                      push eax
    'Reference to '__except_list'
'00722352    64892500000000          mov dword ptr fs:[00000000], esp
'00722359    83ec3c                  sub esp, 3c
'0072235c    53                      push ebx
'0072235d    56                      push esi
'0072235e    57                      push edi
'0072235f    8965f4                  mov dword ptr [ebp-0c], esp
'00722362    c745f890704000          mov dword ptr [ebp-08], 00407090
'00722369    33db                    xor ebx, ebx
'0072236b    895dfc                  mov dword ptr [ebp-04], ebx
'0072236e    8b7508                  mov esi, dword ptr [ebp+08]
'00722371    8b06                    mov eax, dword ptr [esi]
'00722373    56                      push esi
'00722374    ff5004                  call dword ptr [eax+04]
'00722377    8b4d10                  mov ecx, dword ptr [ebp+10]
'0072237a    8b550c                  mov edx, dword ptr [ebp+0c]
'0072237d    8919                    mov dword ptr [ecx], ebx
'0072237f    8b02                    mov eax, dword ptr [edx]
'00722381    3d88130000              cmp eax, 00001388
'00722386    b901000000              mov ecx, 00000001
'0072238b    895ddc                  mov dword ptr [ebp-24], ebx
'0072238e    895dd8                  mov dword ptr [ebp-28], ebx
'00722391    895dd4                  mov dword ptr [ebp-2c], ebx
'00722394    895dd0                  mov dword ptr [ebp-30], ebx
'00722397    895dcc                  mov dword ptr [ebp-34], ebx
'0072239a    895dc8                  mov dword ptr [ebp-38], ebx
'0072239d    895dc4                  mov dword ptr [ebp-3c], ebx
'007223a0    894de0                  mov dword ptr [ebp-20], ecx
'007223a3    7e07                    jle 7223ac

If (arg_0 > 5000) Then
'007223a5    c745e002000000          mov dword ptr [ebp-20], 00000002
    
End If
'007223ac    3de02e0000              cmp eax, 00002ee0
'007223b1    7e10                    jle 7223c3

If (arg_0 > 12000) Then
'007223b3    668b55e0                mov dx, word ptr [ebp-20]
'007223b7    6603d1                  add dx, cx
    var_num4 = var_3 + 1
'007223ba    0f8029040000            jo 7227e9
'007223c0    8955e0                  mov dword ptr [ebp-20], edx
    
End If
'007223c3    3da8610000              cmp eax, 000061a8
'007223c8    7e10                    jle 7223da

If (arg_0 > 25000) Then
'007223ca    668b45e0                mov ax, word ptr [ebp-20]
'007223ce    6603c1                  add ax, cx
    var_num1 = var_num4 + 1
'007223d1    0f8012040000            jo 7227e9
'007223d7    8945e0                  mov dword ptr [ebp-20], eax
    
End If

' *** Reference to "__vbaHresultCheckObj"
'007223da    8b3d80104000            mov edi, dword ptr [00401080]
'007223e0    894de8                  mov dword ptr [ebp-18], ecx
'007223e3    668b4de0                mov cx, word ptr [ebp-20]
'007223e7    66394de8                cmp word ptr [ebp-18], cx
'007223eb    0f8fa6030000            jg 722797

Do While (1 <= var_num1)
'007223f1    8b55dc                  mov edx, dword ptr [ebp-24]
'007223f4    52                      push edx
'007223f5    68cc134300              push 004313cc

' *** Reference to "__vbaStrCmp"
'007223fa    ff1530114000            call dword ptr [00401130]
'00722400    85c0                    test ax, ax
'00722402    741a                    je 72241e
    
    If (    ((vbNullString) <> (vbNullChar))) Then
'00722404    8b45dc                  mov eax, dword ptr [ebp-24]
'00722407    50                      push eax
'00722408    68fc8c4300              push 00438cfc

' *** Reference to "__vbaStrCat"
'0072240d    ff1570104000            call dword ptr [00401070]
    var_49 = (vbNullString) & (", ")
'00722413    8bd0                    mov edx, eax
'00722415    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'00722418    ff15d0124000            call dword ptr [004012d0]
    
End If
'0072241e    8b4d0c                  mov ecx, dword ptr [ebp+0c]
'00722421    833900                  cmp dword ptr [ecx], 00
'00722424    7c43                    jl 722469

If (arg_0 >= 0) Then
'00722426    8b16                    mov edx, dword ptr [esi]
'00722428    8d45c4                  lea eax, dword ptr [ebp-3c]
'0072242b    50                      push eax
'0072242c    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0072242f    51                      push ecx
'00722430    8d45cc                  lea eax, dword ptr [ebp-34]
'00722433    50                      push eax
'00722434    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00722437    51                      push ecx
'00722438    56                      push esi
'00722439    c745c8ffffffff          mov dword ptr [ebp-38], ffffffff
'00722440    c745cc14000000          mov dword ptr [ebp-34], 00000014
'00722447    c745d001000000          mov dword ptr [ebp-30], 00000001
'0072244e    ff9214070000            call dword ptr [edx+00000714]
'00722454    85c0                    test ax, ax
'00722456    7d0e                    jge 722466
'00722458    6814070000              push 00000714
'0072245d    68b0154300              push 004315b0
'00722462    56                      push esi
'00722463    50                      push eax
'00722464    ffd7                    call edi
'00722466    8b5dc4                  mov ebx, dword ptr [ebp-3c]
    
End If
'00722469    8b550c                  mov edx, dword ptr [ebp+0c]
'0072246c    833a50                  cmp dword ptr [edx], 50
'0072246f    7e43                    jle 7224b4

If (arg_0 > 80) Then
'00722471    8b06                    mov eax, dword ptr [esi]
'00722473    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00722476    51                      push ecx
'00722477    8d55c8                  lea edx, dword ptr [ebp-38]
'0072247a    52                      push edx
'0072247b    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0072247e    51                      push ecx
'0072247f    8d55d0                  lea edx, dword ptr [ebp-30]
'00722482    52                      push edx
'00722483    56                      push esi
'00722484    c745c800000000          mov dword ptr [ebp-38], 00000000
'0072248b    c745cc14000000          mov dword ptr [ebp-34], 00000014
'00722492    c745d001000000          mov dword ptr [ebp-30], 00000001
'00722499    ff9014070000            call dword ptr [eax+00000714]
'0072249f    85c0                    test ax, ax
'007224a1    7d0e                    jge 7224b1
'007224a3    6814070000              push 00000714
'007224a8    68b0154300              push 004315b0
'007224ad    56                      push esi
'007224ae    50                      push eax
'007224af    ffd7                    call edi
'007224b1    8b5dc4                  mov ebx, dword ptr [ebp-3c]
    
End If
'007224b4    8b450c                  mov eax, dword ptr [ebp+0c]
'007224b7    813890010000            cmp dword ptr [eax], 00000190
'007224bd    7e40                    jle 7224ff

If (arg_0 > 400) Then
'007224bf    8b0e                    mov ecx, dword ptr [esi]
'007224c1    b801000000              mov eax, 00000001
'007224c6    8945c8                  mov dword ptr [ebp-38], eax
'007224c9    8945d0                  mov dword ptr [ebp-30], eax
'007224cc    8d55c4                  lea edx, dword ptr [ebp-3c]
'007224cf    52                      push edx
'007224d0    8d45c8                  lea eax, dword ptr [ebp-38]
'007224d3    50                      push eax
'007224d4    8d55cc                  lea edx, dword ptr [ebp-34]
'007224d7    52                      push edx
'007224d8    8d45d0                  lea eax, dword ptr [ebp-30]
'007224db    50                      push eax
'007224dc    56                      push esi
'007224dd    c745cc14000000          mov dword ptr [ebp-34], 00000014
'007224e4    ff9114070000            call dword ptr [ecx+00000714]
'007224ea    85c0                    test ax, ax
'007224ec    7d0e                    jge 7224fc
'007224ee    6814070000              push 00000714
'007224f3    68b0154300              push 004315b0
'007224f8    56                      push esi
'007224f9    50                      push eax
'007224fa    ffd7                    call edi
'007224fc    8b5dc4                  mov ebx, dword ptr [ebp-3c]
    
End If
'007224ff    8b4d0c                  mov ecx, dword ptr [ebp+0c]
'00722502    813984030000            cmp dword ptr [ecx], 00000384
'00722508    7e43                    jle 72254d

If (arg_0 > 900) Then
'0072250a    8b16                    mov edx, dword ptr [esi]
'0072250c    8d45c4                  lea eax, dword ptr [ebp-3c]
'0072250f    50                      push eax
'00722510    8d4dc8                  lea ecx, dword ptr [ebp-38]
'00722513    51                      push ecx
'00722514    8d45cc                  lea eax, dword ptr [ebp-34]
'00722517    50                      push eax
'00722518    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0072251b    51                      push ecx
'0072251c    56                      push esi
'0072251d    c745c802000000          mov dword ptr [ebp-38], 00000002
'00722524    c745cc14000000          mov dword ptr [ebp-34], 00000014
'0072252b    c745d001000000          mov dword ptr [ebp-30], 00000001
'00722532    ff9214070000            call dword ptr [edx+00000714]
'00722538    85c0                    test ax, ax
'0072253a    7d0e                    jge 72254a
'0072253c    6814070000              push 00000714
'00722541    68b0154300              push 004315b0
'00722546    56                      push esi
'00722547    50                      push eax
'00722548    ffd7                    call edi
'0072254a    8b5dc4                  mov ebx, dword ptr [ebp-3c]
    
End If
'0072254d    8b550c                  mov edx, dword ptr [ebp+0c]
'00722550    813ad0070000            cmp dword ptr [edx], 000007d0
'00722556    7e43                    jle 72259b

If (arg_0 > 2000) Then
'00722558    8b06                    mov eax, dword ptr [esi]
'0072255a    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0072255d    51                      push ecx
'0072255e    8d55c8                  lea edx, dword ptr [ebp-38]
'00722561    52                      push edx
'00722562    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00722565    51                      push ecx
'00722566    8d55d0                  lea edx, dword ptr [ebp-30]
'00722569    52                      push edx
'0072256a    56                      push esi
'0072256b    c745c803000000          mov dword ptr [ebp-38], 00000003
'00722572    c745cc14000000          mov dword ptr [ebp-34], 00000014
'00722579    c745d001000000          mov dword ptr [ebp-30], 00000001
'00722580    ff9014070000            call dword ptr [eax+00000714]
'00722586    85c0                    test ax, ax
'00722588    7d0e                    jge 722598
'0072258a    6814070000              push 00000714
'0072258f    68b0154300              push 004315b0
'00722594    56                      push esi
'00722595    50                      push eax
'00722596    ffd7                    call edi
'00722598    8b5dc4                  mov ebx, dword ptr [ebp-3c]
    
End If
'0072259b    8b450c                  mov eax, dword ptr [ebp+0c]
'0072259e    813888130000            cmp dword ptr [eax], 00001388
'007225a4    7e43                    jle 7225e9

If (arg_0 > 5000) Then
'007225a6    8b0e                    mov ecx, dword ptr [esi]
'007225a8    8d55c4                  lea edx, dword ptr [ebp-3c]
'007225ab    52                      push edx
'007225ac    8d45c8                  lea eax, dword ptr [ebp-38]
'007225af    50                      push eax
'007225b0    8d55cc                  lea edx, dword ptr [ebp-34]
'007225b3    52                      push edx
'007225b4    8d45d0                  lea eax, dword ptr [ebp-30]
'007225b7    50                      push eax
'007225b8    56                      push esi
'007225b9    c745c804000000          mov dword ptr [ebp-38], 00000004
'007225c0    c745cc14000000          mov dword ptr [ebp-34], 00000014
'007225c7    c745d001000000          mov dword ptr [ebp-30], 00000001
'007225ce    ff9114070000            call dword ptr [ecx+00000714]
'007225d4    85c0                    test ax, ax
'007225d6    7d0e                    jge 7225e6
'007225d8    6814070000              push 00000714
'007225dd    68b0154300              push 004315b0
'007225e2    56                      push esi
'007225e3    50                      push eax
'007225e4    ffd7                    call edi
'007225e6    8b5dc4                  mov ebx, dword ptr [ebp-3c]
    
End If
'007225e9    8b4d0c                  mov ecx, dword ptr [ebp+0c]
'007225ec    8139e02e0000            cmp dword ptr [ecx], 00002ee0
'007225f2    7e43                    jle 722637

If (arg_0 > 12000) Then
'007225f4    8b16                    mov edx, dword ptr [esi]
'007225f6    8d45c4                  lea eax, dword ptr [ebp-3c]
'007225f9    50                      push eax
'007225fa    8d4dc8                  lea ecx, dword ptr [ebp-38]
'007225fd    51                      push ecx
'007225fe    8d45cc                  lea eax, dword ptr [ebp-34]
'00722601    50                      push eax
'00722602    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00722605    51                      push ecx
'00722606    56                      push esi
'00722607    c745c805000000          mov dword ptr [ebp-38], 00000005
'0072260e    c745cc14000000          mov dword ptr [ebp-34], 00000014
'00722615    c745d001000000          mov dword ptr [ebp-30], 00000001
'0072261c    ff9214070000            call dword ptr [edx+00000714]
'00722622    85c0                    test ax, ax
'00722624    7d0e                    jge 722634
'00722626    6814070000              push 00000714
'0072262b    68b0154300              push 004315b0
'00722630    56                      push esi
'00722631    50                      push eax
'00722632    ffd7                    call edi
'00722634    8b5dc4                  mov ebx, dword ptr [ebp-3c]
    
End If
'00722637    8b550c                  mov edx, dword ptr [ebp+0c]
'0072263a    813aa8610000            cmp dword ptr [edx], 000061a8
'00722640    7e43                    jle 722685

If (arg_0 > 25000) Then
'00722642    8b06                    mov eax, dword ptr [esi]
'00722644    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00722647    51                      push ecx
'00722648    8d55c8                  lea edx, dword ptr [ebp-38]
'0072264b    52                      push edx
'0072264c    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0072264f    51                      push ecx
'00722650    8d55d0                  lea edx, dword ptr [ebp-30]
'00722653    52                      push edx
'00722654    56                      push esi
'00722655    c745c806000000          mov dword ptr [ebp-38], 00000006
'0072265c    c745cc14000000          mov dword ptr [ebp-34], 00000014
'00722663    c745d001000000          mov dword ptr [ebp-30], 00000001
'0072266a    ff9014070000            call dword ptr [eax+00000714]
'00722670    85c0                    test ax, ax
'00722672    7d0e                    jge 722682
'00722674    6814070000              push 00000714
'00722679    68b0154300              push 004315b0
'0072267e    56                      push esi
'0072267f    50                      push eax
'00722680    ffd7                    call edi
'00722682    8b5dc4                  mov ebx, dword ptr [ebp-3c]
    
End If
'00722685    6683fb0e                cmp bx, 0e
'00722689    7d6c                    jge 7226f7
'0072268b    8b45dc                  mov eax, dword ptr [ebp-24]
'0072268e    50                      push eax
'0072268f    68b8644500              push 004564b8

' *** Reference to "__vbaStrCat"
'00722694    ff1570104000            call dword ptr [00401070]
var_63 = (var_49) & ("Traditionnel")
'0072269a    8bd0                    mov edx, eax
'0072269c    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'0072269f    ff15d0124000            call dword ptr [004012d0]
'007226a5    8b0e                    mov ecx, dword ptr [esi]
'007226a7    8d55c4                  lea edx, dword ptr [ebp-3c]
'007226aa    52                      push edx
'007226ab    8d45c8                  lea eax, dword ptr [ebp-38]
'007226ae    50                      push eax
'007226af    8d55cc                  lea edx, dword ptr [ebp-34]
'007226b2    52                      push edx
'007226b3    8d45d0                  lea eax, dword ptr [ebp-30]
'007226b6    50                      push eax
'007226b7    56                      push esi
'007226b8    c745c800000000          mov dword ptr [ebp-38], 00000000
'007226bf    c745cc14000000          mov dword ptr [ebp-34], 00000014
'007226c6    c745d001000000          mov dword ptr [ebp-30], 00000001
'007226cd    ff9114070000            call dword ptr [ecx+00000714]
'007226d3    85c0                    test ax, ax
'007226d5    7d0e                    jge 7226e5
'007226d7    6814070000              push 00000714
'007226dc    68b0154300              push 004315b0
'007226e1    56                      push esi
'007226e2    50                      push eax
'007226e3    ffd7                    call edi
'007226e5    66837dc414              cmp word ptr [ebp-3c], 14
'007226ea    7536                    jne 722722

If (0 = 20) Then
'007226ec    8b4ddc                  mov ecx, dword ptr [ebp-24]
'007226ef    51                      push ecx
'007226f0    686c744500              push 0045746c
'007226f5    eb1a                    jmp 722711
'007226f7    6683fb12                cmp bx, 12
'007226fb    7e0b                    jle 722708
'007226fd    8b55dc                  mov edx, dword ptr [ebp-24]
'00722700    52                      push edx
'00722701    68a4744500              push 004574a4
'00722706    eb09                    jmp 722711
'00722708    8b45dc                  mov eax, dword ptr [ebp-24]
'0072270b    50                      push eax
'0072270c    68b8744500              push 004574b8

' *** Reference to "__vbaStrCat"
'00722711    ff1570104000            call dword ptr [00401070]
    var_36 = (var_63) & ("Inhabituel")
'00722717    8bd0                    mov edx, eax
'00722719    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'0072271c    ff15d0124000            call dword ptr [004012d0]
    
End If
'00722722    8b0e                    mov ecx, dword ptr [esi]
'00722724    8d55d8                  lea edx, dword ptr [ebp-28]
'00722727    52                      push edx
'00722728    56                      push esi
'00722729    ff9118070000            call dword ptr [ecx+00000718]
'0072272f    85c0                    test ax, ax
'00722731    7d0e                    jge 722741
'00722733    6818070000              push 00000718
'00722738    68b0154300              push 004315b0
'0072273d    56                      push esi
'0072273e    50                      push eax
'0072273f    ffd7                    call edi
'00722741    8b45dc                  mov eax, dword ptr [ebp-24]
'00722744    50                      push eax
'00722745    68d4744500              push 004574d4

' *** Reference to "__vbaStrCat"
'0072274a    ff1570104000            call dword ptr [00401070]
var_76 = (var_36) & (" d'alignement ")
'00722750    8bd0                    mov edx, eax
'00722752    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaStrMove"
'00722755    ff15d0124000            call dword ptr [004012d0]
'0072275b    8b4dd8                  mov ecx, dword ptr [ebp-28]
'0072275e    50                      push eax
'0072275f    51                      push ecx

' *** Reference to "__vbaStrCat"
'00722760    ff1570104000            call dword ptr [00401070]
var_12 = (var_76) & (vbNullString)
'00722766    8bd0                    mov edx, eax
'00722768    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'0072276b    ff15d0124000            call dword ptr [004012d0]
'00722771    8d55d8                  lea edx, dword ptr [ebp-28]
'00722774    52                      push edx
'00722775    8d45d4                  lea eax, dword ptr [ebp-2c]
'00722778    50                      push eax
'00722779    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'0072277b    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4512, 0)
'00722781    b801000000              mov eax, 00000001
'00722786    83c40c                  add esp, 0c
'00722789    660345e8                add ax, word ptr [ebp-18]
var_num1 = 1 + 1
'0072278d    705a                    jo 7227e9
'0072278f    8945e8                  mov dword ptr [ebp-18], eax
'00722792    e94cfcffff              jmp 7223e3

'ERROR: Two many next close:
Loop
'00722797    68c2277200              push 007227c2
'0072279c    eb23                    jmp 7227c1
'0072279e    f645fc04                test byte ptr [ebp-04], 04
'007227a2    7409                    je 7227ad
'007227a4    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeStr"
'007227a7    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_39)
'007227ad    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'007227b0    51                      push ecx
'007227b1    8d55d8                  lea edx, dword ptr [ebp-28]
'007227b4    52                      push edx
'007227b5    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'007227b7    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4512)
'007227bd    83c40c                  add esp, 0c
'007227c0    c3                      ret
'007227c1    c3                      ret
'007227c2    8b4508                  mov eax, dword ptr [ebp+08]
'007227c5    8b08                    mov ecx, dword ptr [eax]
'007227c7    50                      push eax
'007227c8    ff5108                  call dword ptr [ecx+08]
'007227cb    8b45dc                  mov eax, dword ptr [ebp-24]
'007227ce    8b5510                  mov edx, dword ptr [ebp+10]
'007227d1    8902                    mov dword ptr [edx], eax
'007227d3    8b45fc                  mov eax, dword ptr [ebp-04]
'007227d6    8b4dec                  mov ecx, dword ptr [ebp-14]
'007227d9    5f                      pop edi
'007227da    5e                      pop esi
    'Reference to '__except_list'
'007227db    64890d00000000          mov dword ptr fs:[00000000], ecx
'007227e2    5b                      pop ebx
'007227e3    8be5                    mov esp, ebp
'007227e5    5d                      pop ebp
'007227e6    c20c00                  ret 000c


End Function


Public Function des(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'007227f0    55                      push ebp
'007227f1    8bec                    mov ebp, esp
'007227f3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'007227f6    6866724000              push 00407266
'007227fb    64a100000000            mov ax, word ptr fs:[00000000]
'00722801    50                      push eax
    'Reference to '__except_list'
'00722802    64892500000000          mov dword ptr fs:[00000000], esp
'00722809    83ec4c                  sub esp, 4c
'0072280c    53                      push ebx
'0072280d    56                      push esi
'0072280e    57                      push edi
'0072280f    8965f4                  mov dword ptr [ebp-0c], esp
'00722812    c745f8a0704000          mov dword ptr [ebp-08], 004070a0
'00722819    33ff                    xor edi, edi
'0072281b    897dfc                  mov dword ptr [ebp-04], edi
'0072281e    8b4508                  mov eax, dword ptr [ebp+08]
'00722821    8b08                    mov ecx, dword ptr [eax]
'00722823    50                      push eax
'00722824    ff5104                  call dword ptr [ecx+04]
'00722827    8b550c                  mov edx, dword ptr [ebp+0c]
'0072282a    668b02                  mov ax, word ptr [edx]
'0072282d    bb01000000              mov ebx, 00000001
'00722832    897de4                  mov dword ptr [ebp-1c], edi
'00722835    897dd4                  mov dword ptr [ebp-2c], edi
'00722838    897de4                  mov dword ptr [ebp-1c], edi
'0072283b    668945b8                mov word ptr [ebp-48], ax
'0072283f    8bf3                    mov esi, ebx
'00722841    663b75b8                cmp si, word ptr [ebp-48]
'00722845    7f7e                    jg 7228c5

Do While (1 <= WORD PTR [EBP-48])
'00722847    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0072284a    51                      push ecx
'0072284b    c745dce8030000          mov dword ptr [ebp-24], 000003e8
'00722852    c745d402000000          mov dword ptr [ebp-2c], 00000002

' *** Reference to "rtcRandomNext"
'00722859    ff15a4104000            call dword ptr [004010a4]
'0072285f    d95dc0                  fstp dword ptr [ebp-40]
    'var_5 = (00)
'00722862    8b5510                  mov edx, dword ptr [ebp+10]
'00722865    0fbf02                  movsx eax, word ptr [edx]
'00722868    8945ac                  mov dword ptr [ebp-54], eax
'0072286b    db45ac                  fild dword ptr [ebp-54]
'0072286e    d95da8                  fstp dword ptr [ebp-58]
    'var_86 = (00)
'00722871    d945a8                  fld dword ptr [ebp-58]
'00722874    d84dc0                  fmul dword ptr [ebp-40]
'00722877    dfe0                    fnstsw ax
'00722879    a80d                    test al, 0d
'0072287b    0f858e000000            jne 72290f

' *** Reference to "__vbaFPInt"
'00722881    ff15fc124000            call dword ptr [004012fc]
'00722887    0fbfcf                  movsx ecx, di
'0072288a    894da4                  mov dword ptr [ebp-5c], ecx
'0072288d    db45a4                  fild dword ptr [ebp-5c]
'00722890    d95da0                  fstp dword ptr [ebp-60]
    'var_7 = (00)
'00722893    d845a0                  fadd dword ptr [ebp-60]
'00722896    d80574574000            fadd dword ptr [00405774]
'0072289c    dfe0                    fnstsw ax
'0072289e    a80d                    test al, 0d
'007228a0    756d                    jne 72290f

' *** Reference to "__vbaFpI2"
'007228a2    ff15a0124000            call dword ptr [004012a0]
    var_num1 = CLng(((Int((((arg_1)) * Rnd(var_39))) + (0)) + 0!))
'007228a8    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'007228ab    8bf8                    mov edi, eax
'007228ad    897de4                  mov dword ptr [ebp-1c], edi

' *** Reference to "__vbaFreeVar"
'007228b0    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_44)
'007228b6    668bd3                  mov dx, bx
'007228b9    6603d6                  add dx, si
    var_num4 = 1 + 1
'007228bc    7056                    jo 722914
'007228be    8bf2                    mov esi, edx
'007228c0    e97cffffff              jmp 722841
    
Loop
'007228c5    8b4514                  mov eax, dword ptr [ebp+14]
'007228c8    668b08                  mov cx, word ptr [eax]
'007228cb    6603cf                  add cx, di
var_num3 = arg_2 + 0
'007228ce    7044                    jo 722914
'007228d0    894de4                  mov dword ptr [ebp-1c], ecx
'007228d3    9b                      fwait
'007228d4    68e6287200              push 007228e6
'007228d9    eb0a                    jmp 7228e5
'007228db    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaFreeVar"
'007228de    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_44)
'007228e4    c3                      ret
'007228e5    c3                      ret
'007228e6    8b4508                  mov eax, dword ptr [ebp+08]
'007228e9    8b10                    mov edx, dword ptr [eax]
'007228eb    50                      push eax
'007228ec    ff5208                  call dword ptr [edx+08]
'007228ef    8b4518                  mov eax, dword ptr [ebp+18]
'007228f2    668b4de4                mov cx, word ptr [ebp-1c]
'007228f6    668908                  mov word ptr [eax], cx
'007228f9    8b45fc                  mov eax, dword ptr [ebp-04]
'007228fc    8b4dec                  mov ecx, dword ptr [ebp-14]
'007228ff    5f                      pop edi
'00722900    5e                      pop esi
    'Reference to '__except_list'
'00722901    64890d00000000          mov dword ptr fs:[00000000], ecx
'00722908    5b                      pop ebx
'00722909    8be5                    mov esp, ebp
'0072290b    5d                      pop ebp
'0072290c    c21400                  ret 0014


End Function


Public Function Alignement(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'00722920    55                      push ebp
'00722921    8bec                    mov ebp, esp
'00722923    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00722926    6866724000              push 00407266
'0072292b    64a100000000            mov ax, word ptr fs:[00000000]
'00722931    50                      push eax
    'Reference to '__except_list'
'00722932    64892500000000          mov dword ptr fs:[00000000], esp
'00722939    83ec24                  sub esp, 24
'0072293c    53                      push ebx
'0072293d    56                      push esi
'0072293e    57                      push edi
'0072293f    8965f4                  mov dword ptr [ebp-0c], esp
'00722942    c745f8b0704000          mov dword ptr [ebp-08], 004070b0
'00722949    33ff                    xor edi, edi
'0072294b    897dfc                  mov dword ptr [ebp-04], edi
'0072294e    8b7508                  mov esi, dword ptr [ebp+08]
'00722951    8b06                    mov eax, dword ptr [esi]
'00722953    56                      push esi
'00722954    ff5004                  call dword ptr [eax+04]
'00722957    8b4d0c                  mov ecx, dword ptr [ebp+0c]
'0072295a    8939                    mov dword ptr [ecx], edi
'0072295c    8b16                    mov edx, dword ptr [esi]
'0072295e    8d45d4                  lea eax, dword ptr [ebp-2c]
'00722961    50                      push eax
'00722962    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00722965    51                      push ecx
'00722966    8d45dc                  lea eax, dword ptr [ebp-24]
'00722969    50                      push eax
'0072296a    8d4de0                  lea ecx, dword ptr [ebp-20]
'0072296d    51                      push ecx
'0072296e    897de0                  mov dword ptr [ebp-20], edi
'00722971    897ddc                  mov dword ptr [ebp-24], edi
'00722974    897dd8                  mov dword ptr [ebp-28], edi
'00722977    56                      push esi
'00722978    897de4                  mov dword ptr [ebp-1c], edi
'0072297b    897dd4                  mov dword ptr [ebp-2c], edi
'0072297e    897dd8                  mov dword ptr [ebp-28], edi
'00722981    c745dc64000000          mov dword ptr [ebp-24], 00000064
'00722988    c745e001000000          mov dword ptr [ebp-20], 00000001
'0072298f    ff9214070000            call dword ptr [edx+00000714]
'00722995    3bc7                    cmp eax, edi
'00722997    7d12                    jge 7229ab
'00722999    6814070000              push 00000714
'0072299e    68b0154300              push 004315b0
'007229a3    56                      push esi
'007229a4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007229a5    ff1580104000            call dword ptr [00401080]
'007229ab    8b75d4                  mov esi, dword ptr [ebp-2c]
'007229ae    663bf7                  cmp si, di
'007229b1    7e12                    jle 7229c5

If (0 > 0) Then

' *** Reference to "__vbaStrCopy"
'007229b3    8b3d48124000            mov edi, dword ptr [00401248]
'007229b9    baf8744500              mov edx, 004574f8
'007229be    8d4de4                  lea ecx, dword ptr [ebp-1c]
'007229c1    ffd7                    call edi
    var_40 = ("Loyal bon")
'007229c3    eb06                    jmp 7229cb
    
End If

' *** Reference to "__vbaStrCopy"
'007229c5    8b3d48124000            mov edi, dword ptr [00401248]
'007229cb    6683fe23                cmp si, 23
'007229cf    7e0a                    jle 7229db

If (0 > 35) Then
'007229d1    ba10754500              mov edx, 00457510
'007229d6    8d4de4                  lea ecx, dword ptr [ebp-1c]
'007229d9    ffd7                    call edi
    var_40 = ("Neutre bon")
End If
'007229db    6683fe27                cmp si, 27
'007229df    7e0a                    jle 7229eb

If (0 > 39) Then
'007229e1    ba2c754500              mov edx, 0045752c
'007229e6    8d4de4                  lea ecx, dword ptr [ebp-1c]
'007229e9    ffd7                    call edi
    var_40 = ("Chaotique bon")
End If
'007229eb    6683fe29                cmp si, 29
'007229ef    7e0a                    jle 7229fb

If (0 > 41) Then
'007229f1    ba4c754500              mov edx, 0045754c
'007229f6    8d4de4                  lea ecx, dword ptr [ebp-1c]
'007229f9    ffd7                    call edi
    var_40 = ("Loyal neutre")
End If
'007229fb    6683fe3d                cmp si, 3d
'007229ff    7e0a                    jle 722a0b

If (0 > 61) Then
'00722a01    ba887b4300              mov edx, 00437b88
'00722a06    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00722a09    ffd7                    call edi
    var_40 = ("Neutre")
End If
'00722a0b    6683fe3f                cmp si, 3f
'00722a0f    7e0a                    jle 722a1b

If (0 > 63) Then
'00722a11    ba6c754500              mov edx, 0045756c
'00722a16    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00722a19    ffd7                    call edi
    var_40 = ("Chaotique neutre")
End If
'00722a1b    6683fe40                cmp si, 40
'00722a1f    7e0a                    jle 722a2b

If (0 > 64) Then
'00722a21    ba94754500              mov edx, 00457594
'00722a26    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00722a29    ffd7                    call edi
    var_40 = ("Loyal mauvais")
End If
'00722a2b    6683fe5a                cmp si, 5a
'00722a2f    7e0a                    jle 722a3b

If (0 > 90) Then
'00722a31    bab4754500              mov edx, 004575b4
'00722a36    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00722a39    ffd7                    call edi
    var_40 = ("Neutre mauvais")
End If
'00722a3b    6683fe62                cmp si, 62
'00722a3f    7e0a                    jle 722a4b

If (0 > 98) Then
'00722a41    bad8754500              mov edx, 004575d8
'00722a46    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00722a49    ffd7                    call edi
    var_40 = ("Chaotique mauvais")
End If
'00722a4b    685d2a7200              push 00722a5d
'00722a50    eb0a                    jmp 722a5c
'00722a52    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'00722a55    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'00722a5b    c3                      ret
'00722a5c    c3                      ret
'00722a5d    8b4508                  mov eax, dword ptr [ebp+08]
'00722a60    8b10                    mov edx, dword ptr [eax]
'00722a62    50                      push eax
'00722a63    ff5208                  call dword ptr [edx+08]
'00722a66    8b450c                  mov eax, dword ptr [ebp+0c]
'00722a69    8b4de4                  mov ecx, dword ptr [ebp-1c]
'00722a6c    8908                    mov dword ptr [eax], ecx
'00722a6e    8b45fc                  mov eax, dword ptr [ebp-04]
'00722a71    8b4dec                  mov ecx, dword ptr [ebp-14]
'00722a74    5f                      pop edi
'00722a75    5e                      pop esi
    'Reference to '__except_list'
'00722a76    64890d00000000          mov dword ptr fs:[00000000], ecx
'00722a7d    5b                      pop ebx
'00722a7e    8be5                    mov esp, ebp
'00722a80    5d                      pop ebp
'00722a81    c20800                  ret 0008


End Function


Public Function PNJ(arg_0 As Unknow, arg_1 As String, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'00722a90    55                      push ebp
'00722a91    8bec                    mov ebp, esp
'00722a93    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00722a96    6866724000              push 00407266
'00722a9b    64a100000000            mov ax, word ptr fs:[00000000]
'00722aa1    50                      push eax
    'Reference to '__except_list'
'00722aa2    64892500000000          mov dword ptr fs:[00000000], esp
'00722aa9    81ec68010000            sub esp, 00000168
'00722aaf    53                      push ebx
'00722ab0    56                      push esi
'00722ab1    57                      push edi
'00722ab2    8965f4                  mov dword ptr [ebp-0c], esp
'00722ab5    c745f8f0704000          mov dword ptr [ebp-08], 004070f0
'00722abc    33f6                    xor esi, esi
'00722abe    8975fc                  mov dword ptr [ebp-04], esi
'00722ac1    8b4508                  mov eax, dword ptr [ebp+08]
'00722ac4    8b08                    mov ecx, dword ptr [eax]
'00722ac6    50                      push eax
'00722ac7    ff5104                  call dword ptr [ecx+04]
'00722aca    8b5510                  mov edx, dword ptr [ebp+10]
'00722acd    56                      push esi
'00722ace    68f0784500              push 004578f0
'00722ad3    8d45bc                  lea eax, dword ptr [ebp-44]
'00722ad6    50                      push eax
'00722ad7    8975e4                  mov dword ptr [ebp-1c], esi
'00722ada    8975d4                  mov dword ptr [ebp-2c], esi
'00722add    8975ac                  mov dword ptr [ebp-54], esi
'00722ae0    8975a8                  mov dword ptr [ebp-58], esi
'00722ae3    8975a4                  mov dword ptr [ebp-5c], esi
'00722ae6    897590                  mov dword ptr [ebp-70], esi
'00722ae9    89758c                  mov dword ptr [ebp-74], esi
'00722aec    897588                  mov dword ptr [ebp-78], esi
'00722aef    897584                  mov dword ptr [ebp-7c], esi
'00722af2    897580                  mov dword ptr [ebp-80], esi
'00722af5    89b570ffffff            mov dword ptr [ebp+ffffff70], esi
'00722afb    89b560ffffff            mov dword ptr [ebp+ffffff60], esi
'00722b01    89b53cffffff            mov dword ptr [ebp+ffffff3c], esi
'00722b07    89b538ffffff            mov dword ptr [ebp+ffffff38], esi
'00722b0d    89b534ffffff            mov dword ptr [ebp+ffffff34], esi
'00722b13    89b530ffffff            mov dword ptr [ebp+ffffff30], esi
'00722b19    89b52cffffff            mov dword ptr [ebp+ffffff2c], esi
'00722b1f    8932                    mov dword ptr [edx], esi

' *** Reference to "__vbaAryConstruct2"
'00722b21    ff1538114000            call dword ptr [00401138]
Dim var_58 (0 To 16) As Unknow
'00722b27    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'00722b2d    51                      push ecx
'00722b2e    c78578ffffffe8030000    mov dword ptr [ebp+ffffff78], 000003e8
'00722b38    c78570ffffff02000000    mov dword ptr [ebp+ffffff70], 00000002

' *** Reference to "rtcRandomNext"
'00722b42    ff15a4104000            call dword ptr [004010a4]
'00722b48    d80de8704000            fmul dword ptr [004070e8]
'00722b4e    dfe0                    fnstsw ax
'00722b50    a80d                    test al, 0d
'00722b52    0f85930f0000            jne 723aeb

' *** Reference to "__vbaFPInt"
'00722b58    ff15fc124000            call dword ptr [004012fc]
'00722b5e    d80574574000            fadd dword ptr [00405774]
'00722b64    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'00722b6a    dd5d9c                  fstp qword ptr [ebp-64]
'var_51 = (00)
'00722b6d    dfe0                    fnstsw ax
'00722b6f    a80d                    test al, 0d
'00722b71    0f85740f0000            jne 723aeb

' *** Reference to "__vbaFreeVar"
'00722b77    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_19)
'00722b7d    dd459c                  fld qword ptr [ebp-64]
'00722b80    dc1de0704000            fcomp qword ptr [004070e0]
'00722b86    dfe0                    fnstsw ax
'00722b88    f6c401                  test ah, 01
'00722b8b    7410                    je 722b9d

If ((61# > ((Int((Rnd(var_87) * 0!)) + 0!)))) Then
'00722b8d    c7459c33333333          mov dword ptr [ebp-64], 33333333
'00722b94    c745a033332040          mov dword ptr [ebp-60], 40203333
'00722b9b    eb2e                    jmp 722bcb
    
End If
'00722b9d    dd459c                  fld qword ptr [ebp-64]
'00722ba0    dc1dd8704000            fcomp qword ptr [004070d8]
'00722ba6    dfe0                    fnstsw ax
'00722ba8    f6c401                  test ah, 01
'00722bab    7410                    je 722bbd

If ((81# > (var_51))) Then
'00722bad    c7459ccdcccccc          mov dword ptr [ebp-64], cccccccd
'00722bb4    c745a0cccc1c40          mov dword ptr [ebp-60], 401ccccc
'00722bbb    eb0e                    jmp 722bcb
    
End If
'00722bbd    c7459c66666666          mov dword ptr [ebp-64], 66666666
'00722bc4    c745a066661c40          mov dword ptr [ebp-60], 401c6666
'00722bcb    8b4dc8                  mov ecx, dword ptr [ebp-38]

' *** Reference to "__vbaStrCopy"
'00722bce    8b3548124000            mov esi, dword ptr [00401248]
'00722bd4    ba80334300              mov edx, 00433380
'00722bd9    ffd6                    call esi
var_84 = ("Adepte")
'00722bdb    8b55c8                  mov edx, dword ptr [ebp-38]
'00722bde    bb01000000              mov ebx, 00000001
'00722be3    66895a04                mov word ptr [edx+04], bx
'00722be7    8b45c8                  mov eax, dword ptr [ebp-38]
'00722bea    bf06000000              mov edi, 00000006
'00722bef    66897806                mov word ptr [eax+06], di
'00722bf3    8b4dc8                  mov ecx, dword ptr [ebp-38]
'00722bf6    baf4344400              mov edx, 004434f4
'00722bfb    83c108                  add ecx, 08
var_num3 = @[(var_58[((2~))]]
'00722bfe    ffd6                    call esi
var_85 = ("Barbare")
'00722c00    8b55c8                  mov edx, dword ptr [ebp-38]
'00722c03    66895a0c                mov word ptr [edx+0c], bx
'00722c07    8b45c8                  mov eax, dword ptr [ebp-38]
'00722c0a    bb04000000              mov ebx, 00000004
'00722c0f    6689580e                mov word ptr [eax+0e], bx
'00722c13    8b4dc8                  mov ecx, dword ptr [ebp-38]
'00722c16    ba80344300              mov edx, 00433480
'00722c1b    83c110                  add ecx, 10
var_num3 = @[(var_58[((4~))]]
'00722c1e    ffd6                    call esi
var_844 = ("Barde")
'00722c20    8b55c8                  mov edx, dword ptr [ebp-38]
'00722c23    66c742140100            mov word ptr [edx+14], 0001
@[(var_58[((20~))]] = 1
'00722c29    8b45c8                  mov eax, dword ptr [ebp-38]
'00722c2c    66897816                mov word ptr [eax+16], di
'00722c30    8b4dc8                  mov ecx, dword ptr [ebp-38]
'00722c33    baf4334300              mov edx, 004333f4
'00722c38    83c118                  add ecx, 18
var_num3 = @[(var_58[((6~))]]
'00722c3b    ffd6                    call esi
var_1180 = ("Druide")
'00722c3d    8b55c8                  mov edx, dword ptr [ebp-38]
'00722c40    66c7421c0100            mov word ptr [edx+1c], 0001
@[(var_58[((28~))]] = 1
'00722c46    8b45c8                  mov eax, dword ptr [ebp-38]
'00722c49    6689781e                mov word ptr [eax+1e], di
'00722c4d    8b4dc8                  mov ecx, dword ptr [ebp-38]
'00722c50    baec9a4300              mov edx, 00439aec
'00722c55    83c120                  add ecx, 20
var_num3 = @[(var_58[((8~))]]
'00722c58    ffd6                    call esi
var_2360 = ("Ensorceleur")
'00722c5a    8b55c8                  mov edx, dword ptr [ebp-38]
'00722c5d    66c742240100            mov word ptr [edx+24], 0001
@[(var_58[((36~))]] = 1
'00722c63    8b45c8                  mov eax, dword ptr [ebp-38]
'00722c66    66895826                mov word ptr [eax+26], bx
'00722c6a    8b4dc8                  mov ecx, dword ptr [ebp-38]
'00722c6d    ba00764500              mov edx, 00457600
'00722c72    83c128                  add ecx, 28
var_num3 = @[(var_58[((10~))]]
'00722c75    ffd6                    call esi
var_2585 = ("Expert")
'00722c77    8b55c8                  mov edx, dword ptr [ebp-38]
'00722c7a    66c7422c0300            mov word ptr [edx+2c], 0003
@[(var_58[((44~))]] = 3
'00722c80    8b45c8                  mov eax, dword ptr [ebp-38]
'00722c83    6689582e                mov word ptr [eax+2e], bx
'00722c87    8b4dc8                  mov ecx, dword ptr [ebp-38]
'00722c8a    bad4624500              mov edx, 004562d4
'00722c8f    83c130                  add ecx, 30
var_num3 = @[(var_58[((12~))]]
'00722c92    ffd6                    call esi
var_2586 = ("Gens du peuple")
'00722c94    8b55c8                  mov edx, dword ptr [ebp-38]
'00722c97    66895a34                mov word ptr [edx+34], bx
'00722c9b    8b45c8                  mov eax, dword ptr [ebp-38]
'00722c9e    66895836                mov word ptr [eax+36], bx
'00722ca2    8b4dc8                  mov ecx, dword ptr [ebp-38]
'00722ca5    ba90974300              mov edx, 00439790
'00722caa    83c138                  add ecx, 38
var_num3 = @[(var_58[((14~))]]
'00722cad    ffd6                    call esi
var_138 = ("Guerrier")
'00722caf    8b55c8                  mov edx, dword ptr [ebp-38]
'00722cb2    66c7423c0100            mov word ptr [edx+3c], 0001
@[(var_58[((60~))]] = 1
'00722cb8    8b45c8                  mov eax, dword ptr [ebp-38]
'00722cbb    66c7403e0800            mov word ptr [eax+3e], 0008
@[(var_58[((62~))]] = 8
'00722cc1    8b4dc8                  mov ecx, dword ptr [ebp-38]
'00722cc4    ba14764500              mov edx, 00457614
'00722cc9    83c140                  add ecx, 40
var_num3 = @[(var_58[((16~))]]
'00722ccc    ffd6                    call esi
var_708 = ("Homme d'armes")
'00722cce    8b55c8                  mov edx, dword ptr [ebp-38]
'00722cd1    66c742440200            mov word ptr [edx+44], 0002
@[(var_58[((68~))]] = 2
'00722cd7    8b45c8                  mov eax, dword ptr [ebp-38]
'00722cda    66895846                mov word ptr [eax+46], bx
'00722cde    8b4dc8                  mov ecx, dword ptr [ebp-38]
'00722ce1    bac0eb4300              mov edx, 0043ebc0
'00722ce6    83c148                  add ecx, 48
var_num3 = @[(var_58[((18~))]]
'00722ce9    ffd6                    call esi
var_570 = ("Magicien")
'00722ceb    8b55c8                  mov edx, dword ptr [ebp-38]
'00722cee    66c7424c0100            mov word ptr [edx+4c], 0001
@[(var_58[((76~))]] = 1
'00722cf4    8b45c8                  mov eax, dword ptr [ebp-38]
'00722cf7    6689584e                mov word ptr [eax+4e], bx
'00722cfb    8b4dc8                  mov ecx, dword ptr [ebp-38]
'00722cfe    bacc864300              mov edx, 004386cc
'00722d03    83c150                  add ecx, 50
var_num3 = @[(var_58[((20~))]]
'00722d06    ffd6                    call esi
var_920 = ("Moine")
'00722d08    8b55c8                  mov edx, dword ptr [ebp-38]
'00722d0b    66c742540100            mov word ptr [edx+54], 0001
@[(var_58[((84~))]] = 1
'00722d11    8b45c8                  mov eax, dword ptr [ebp-38]
'00722d14    66895856                mov word ptr [eax+56], bx
'00722d18    8b4dc8                  mov ecx, dword ptr [ebp-38]
'00722d1b    bab86f4500              mov edx, 00456fb8
'00722d20    83c158                  add ecx, 58
var_num3 = @[(var_58[((22~))]]
'00722d23    ffd6                    call esi
var_1118 = ("Noble")
'00722d25    8b55c8                  mov edx, dword ptr [ebp-38]
'00722d28    66c7425c0100            mov word ptr [edx+5c], 0001
@[(var_58[((92~))]] = 1
'00722d2e    8b45c8                  mov eax, dword ptr [ebp-38]
'00722d31    66895866                mov word ptr [eax+66], bx
'00722d35    8b4dc8                  mov ecx, dword ptr [ebp-38]
'00722d38    bacc334300              mov edx, 004333cc
'00722d3d    83c160                  add ecx, 60
var_num3 = @[(var_58[((24~))]]
'00722d40    ffd6                    call esi
var_916 = ("Paladin")
'00722d42    8b55c8                  mov edx, dword ptr [ebp-38]
'00722d45    bb01000000              mov ebx, 00000001
'00722d4a    66895a64                mov word ptr [edx+64], bx
'00722d4e    8b45c8                  mov eax, dword ptr [ebp-38]
'00722d51    66c7406e0300            mov word ptr [eax+6e], 0003
@[(var_58[((110~))]] = 3
'00722d57    8b4dc8                  mov ecx, dword ptr [ebp-38]
'00722d5a    bab8334300              mov edx, 004333b8
'00722d5f    83c168                  add ecx, 68
var_num3 = @[(var_58[((26~))]]
'00722d62    ffd6                    call esi
var_1190 = ("Prêtre")
'00722d64    8b55c8                  mov edx, dword ptr [ebp-38]
'00722d67    66895a6c                mov word ptr [edx+6c], bx
'00722d6b    8b45c8                  mov eax, dword ptr [ebp-38]
'00722d6e    66897876                mov word ptr [eax+76], di
'00722d72    8b4dc8                  mov ecx, dword ptr [ebp-38]
'00722d75    baa8594500              mov edx, 004559a8
'00722d7a    83c170                  add ecx, 70
var_num3 = @[(var_58[((28~))]]
'00722d7d    ffd6                    call esi
var_2353 = ("Rôdeur")
'00722d7f    8b55c8                  mov edx, dword ptr [ebp-38]
'00722d82    66895a74                mov word ptr [edx+74], bx
'00722d86    8b45c8                  mov eax, dword ptr [ebp-38]
'00722d89    66c7407e0300            mov word ptr [eax+7e], 0003
@[(var_58[((126~))]] = 3
'00722d8f    8b4dc8                  mov ecx, dword ptr [ebp-38]
'00722d92    baa8974300              mov edx, 004397a8
'00722d97    83c178                  add ecx, 78
var_num3 = @[(var_58[((30~))]]
'00722d9a    ffd6                    call esi
var_2355 = ("Roublard")
'00722d9c    8b55c8                  mov edx, dword ptr [ebp-38]
'00722d9f    66895a7c                mov word ptr [edx+7c], bx
'00722da3    8b45c8                  mov eax, dword ptr [ebp-38]
'00722da6    bacc134300              mov edx, 004313cc
'00722dab    8d4da8                  lea ecx, dword ptr [ebp-58]
'00722dae    66c780860000000800      mov word ptr [eax+00000086], 0008
@[(var_58[((134~))]] = 8
'00722db7    ffd6                    call esi
var_86 = (vbNullChar)
'00722db9    8b4d0c                  mov ecx, dword ptr [ebp+0c]
'00722dbc    8b01                    mov eax, dword ptr [ecx]
'00722dbe    85c0                    test ax, ax
'00722dc0    7c07                    jl 722dc9
'00722dc2    c745e4fdffffff          mov dword ptr [ebp-1c], fffffffd
'00722dc9    83f850                  cmp eax, 50
'00722dcc    7e07                    jle 722dd5

If (1 > 80) Then
'00722dce    c745e4feffffff          mov dword ptr [ebp-1c], fffffffe
    
End If
'00722dd5    3d90010000              cmp eax, 00000190
'00722dda    7e07                    jle 722de3

If (1 > 400) Then
'00722ddc    c745e4ffffffff          mov dword ptr [ebp-1c], ffffffff
    
End If
'00722de3    3d84030000              cmp eax, 00000384
'00722de8    7e07                    jle 722df1

If (1 > 900) Then
'00722dea    33f6                    xor esi, esi
    var_num8 = Empty
'00722dec    8975e4                  mov dword ptr [ebp-1c], esi
'00722def    eb02                    jmp 722df3
    
End If
'00722df1    33f6                    xor esi, esi
var_num8 = Empty
'00722df3    3dd0070000              cmp eax, 000007d0
'00722df8    7e07                    jle 722e01

If (1 > 2000) Then
'00722dfa    c745e403000000          mov dword ptr [ebp-1c], 00000003
    
End If
'00722e01    3d88130000              cmp eax, 00001388
'00722e06    7e03                    jle 722e0b

If (1 > 5000) Then
'00722e08    897de4                  mov dword ptr [ebp-1c], edi
    
End If
'00722e0b    3de02e0000              cmp eax, 00002ee0
'00722e10    7e07                    jle 722e19

If (1 > 12000) Then
'00722e12    c745e409000000          mov dword ptr [ebp-1c], 00000009
    
End If
'00722e19    3da8610000              cmp eax, 000061a8
'00722e1e    7e07                    jle 722e27

If (1 > 25000) Then
'00722e20    c745e40c000000          mov dword ptr [ebp-1c], 0000000c
    
End If
'00722e27    0fbf55e4                movsx edx, word ptr [ebp-1c]
'00722e2b    8995e4feffff            mov dword ptr [ebp+fffffee4], edx
'00722e31    db85e4feffff            fild dword ptr [ebp+fffffee4]
'00722e37    dd9ddcfeffff            fstp qword ptr [ebp+fffffedc]
'var_10 = (00)
'00722e3d    dd85dcfeffff            fld qword ptr [ebp+fffffedc]
'00722e43    833d00b0720000          cmp dword ptr [0072b000], 00
'00722e4a    7508                    jne 722e54

If (vbNullChar = 0) Then
'00722e4c    dc3548164000            fdiv qword ptr [00401648]
'00722e52    eb11                    jmp 722e65
    
End If
'00722e54    ff354c164000            push dword ptr [0040164c]
'00722e5a    ff3548164000            push dword ptr [00401648]

' *** Reference to "_adj_fdiv_m64"
'00722e60    e81f44ceff              call 407284
'00722e65    dfe0                    fnstsw ax
'00722e67    a80d                    test al, 0d
'00722e69    0f857c0c0000            jne 723aeb

' *** Reference to "__vbaFPInt"
'00722e6f    ff15fc124000            call dword ptr [004012fc]
'00722e75    dd9d68ffffff            fstp qword ptr [ebp+ffffff68]
'var_132 = (00)
'00722e7b    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'00722e81    50                      push eax
'00722e82    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'00722e88    51                      push ecx
'00722e89    c78560ffffff05000000    mov dword ptr [ebp+ffffff60], 00000005
'00722e93    c78578ffffff01000000    mov dword ptr [ebp+ffffff78], 00000001
'00722e9d    c78570ffffff02000000    mov dword ptr [ebp+ffffff70], 00000002
'00722ea7    e844f8dcff              call 4f26f0
Call sub_4F26F0()
'00722eac    8d9560ffffff            lea edx, dword ptr [ebp+ffffff60]
'00722eb2    52                      push edx
'00722eb3    8945b0                  mov dword ptr [ebp-50], eax
'00722eb6    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'00722ebc    50                      push eax
'00722ebd    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'00722ebf    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_19, var_20)

' *** Reference to "__vbaStrMove"
'00722ec5    8b1dd0124000            mov ebx, dword ptr [004012d0]

' *** Reference to "__vbaStrCat"
'00722ecb    8b3d70104000            mov edi, dword ptr [00401070]
'00722ed1    8975e0                  mov dword ptr [ebp-20], esi
'00722ed4    8b45e0                  mov eax, dword ptr [ebp-20]
'00722ed7    897598                  mov dword ptr [ebp-68], esi
'00722eda    8b7508                  mov esi, dword ptr [ebp+08]
'00722edd    83c40c                  add esp, 0c
'00722ee0    b90f000000              mov ecx, 0000000f
'00722ee5    663bc1                  cmp ax, cx
'00722ee8    0f8f84060000            jg 723572

Do While (__vbaStrCopy <= 15)
'00722eee    663d0300                cmp ax, 0003
'00722ef2    c745d800000000          mov dword ptr [ebp-28], 00000000
'00722ef9    7406                    je 722f01
    
    If (    __vbaStrCopy <> 3) Then
'00722efb    663d0e00                cmp ax, 000e
'00722eff    7571                    jne 722f72
    
    If (    __vbaStrCopy = 14) Then
'00722f01    66837de4ff              cmp word ptr [ebp-1c], ffffffff
'00722f06    7d6a                    jge 722f72
    
    If (    var_40 < -1) Then
'00722f08    8b0e                    mov ecx, dword ptr [esi]
'00722f0a    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'00722f10    52                      push edx
'00722f11    8d8534ffffff            lea eax, dword ptr [ebp+ffffff34]
'00722f17    50                      push eax
'00722f18    8d9538ffffff            lea edx, dword ptr [ebp+ffffff38]
'00722f1e    52                      push edx
'00722f1f    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'00722f25    50                      push eax
'00722f26    56                      push esi
'00722f27    c78534ffffff00000000    mov dword ptr [ebp+ffffff34], 00000000
'00722f31    c78538ffffff14000000    mov dword ptr [ebp+ffffff38], 00000014
'00722f3b    c7853cffffff01000000    mov dword ptr [ebp+ffffff3c], 00000001
'00722f45    ff9114070000            call dword ptr [ecx+00000714]
'00722f4b    85c0                    test ax, ax
'00722f4d    7d12                    jge 722f61
'00722f4f    6814070000              push 00000714
'00722f54    68b0154300              push 004315b0
'00722f59    56                      push esi
'00722f5a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00722f5b    ff1580104000            call dword ptr [00401080]
'00722f61    6683bd30ffffff14        cmp word ptr [ebp+ffffff30], 14
'00722f69    7507                    jne 722f72
    
    If (    0 = 20) Then
'00722f6b    c745d80a000000          mov dword ptr [ebp-28], 0000000a
    
End If
'00722f72    6a00                    push 00
'00722f74    6a1c                    push 1c
'00722f76    6a01                    push 01
'00722f78    6a02                    push 02
'00722f7a    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00722f7d    51                      push ecx
'00722f7e    6a02                    push 02
'00722f80    6880000000              push 00000080

' *** Reference to "__vbaRedim"
'00722f85    ff1584114000            call dword ptr [00401184]
Dim var_44() As Integer
ReDim var_44(0 To 28)
'00722f8b    83c41c                  add esp, 1c
'00722f8e    c745e801000000          mov dword ptr [ebp-18], 00000001
'00722f95    668b55e8                mov dx, word ptr [ebp-18]
'00722f99    663b55b0                cmp dx, word ptr [ebp-50]
'00722f9d    0f8f2e030000            jg 7232d1

Do While (var_41 <= WORD PTR [EBP-50])
'00722fa3    668b45d8                mov ax, word ptr [ebp-28]
'00722fa7    660345e4                add ax, word ptr [ebp-1c]
    var_num1 = var_45 + var_40
'00722fab    0fbf75e0                movsx esi, word ptr [ebp-20]
'00722faf    0f803b0b0000            jo 723af0
'00722fb5    83fe11                  cmp esi, 11
'00722fb8    89853cffffff            mov dword ptr [ebp+ffffff3c], eax
'00722fbe    7211                    jc 722fd1
    
    If (    __vbaStrCopy >= 17) Then

' *** Reference to "__vbaGenerateBoundsError"
'00722fc0    ff1528114000            call dword ptr [00401128]
    Err.Raise 9
'00722fc6    83fe11                  cmp esi, 11
'00722fc9    7206                    jc 722fd1
    
    If (    __vbaStrCopy >= 17) Then

' *** Reference to "__vbaGenerateBoundsError"
'00722fcb    ff1528114000            call dword ptr [00401128]
    Err.Raise 9
End If
'00722fd1    8b7d08                  mov edi, dword ptr [ebp+08]
'00722fd4    8b0f                    mov ecx, dword ptr [edi]
'00722fd6    8d9538ffffff            lea edx, dword ptr [ebp+ffffff38]
'00722fdc    52                      push edx
'00722fdd    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'00722fe3    50                      push eax
'00722fe4    8b45c8                  mov eax, dword ptr [ebp-38]
'00722fe7    8d54f006                lea edx, dword ptr [esp+06]
'00722feb    52                      push edx
'00722fec    8d44f004                lea eax, dword ptr [esp+04]
'00722ff0    50                      push eax
'00722ff1    57                      push edi
'00722ff2    ff9114070000            call dword ptr [ecx+00000714]
'00722ff8    85c0                    test ax, ax
'00722ffa    7d12                    jge 72300e
'00722ffc    6814070000              push 00000714
'00723001    68b0154300              push 004315b0
'00723006    57                      push edi
'00723007    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00723008    ff1580104000            call dword ptr [00401080]
'0072300e    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'00723014    6685c0                  test eax, eax
'00723017    8945ac                  mov dword ptr [ebp-54], eax
'0072301a    0f8e94020000            jle 7232b4
'00723020    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00723023    85c9                    test cx, cx
'00723025    7429                    je 723050
'00723027    66833901                cmp word ptr [ecx], 01
'0072302b    7523                    jne 723050
'0072302d    8b5114                  mov edx, dword ptr [ecx+14]
'00723030    0fbff8                  movsx edi, ax
'00723033    8b4110                  mov eax, dword ptr [ecx+10]
'00723036    2bfa                    sub edi, edx
var_num7 = var_135 - LBound(var_44)
'00723038    3bf8                    cmp edi, eax
'0072303a    7209                    jc 723045

If (var_num7 >= (UBound(var_44) - LBound(var_44))) Then

' *** Reference to "__vbaGenerateBoundsError"
'0072303c    ff1528114000            call dword ptr [00401128]
    Err.Raise 9
'00723042    8b4dd4                  mov ecx, dword ptr [ebp-2c]
    
End If
'00723045    8d143f                  lea edx, dword ptr [edi+edi]
'00723048    8995d8feffff            mov dword ptr [ebp+fffffed8], edx
'0072304e    eb0f                    jmp 72305f

' *** Reference to "__vbaGenerateBoundsError"
'00723050    ff1528114000            call dword ptr [00401128]
Err.Raise 9
'00723056    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00723059    8985d8feffff            mov dword ptr [ebp+fffffed8], eax
'0072305f    85c9                    test cx, cx
'00723061    7424                    je 723087
'00723063    66833901                cmp word ptr [ecx], 01
'00723067    751e                    jne 723087
'00723069    0fbf7dac                movsx edi, word ptr [ebp-54]
'0072306d    8b5114                  mov edx, dword ptr [ecx+14]
'00723070    8b4110                  mov eax, dword ptr [ecx+10]
'00723073    2bfa                    sub edi, edx
var_num7 = var_135 - LBound(var_44)
'00723075    3bf8                    cmp edi, eax
'00723077    7209                    jc 723082

If (var_num7 >= (UBound(var_44) - LBound(var_44))) Then

' *** Reference to "__vbaGenerateBoundsError"
'00723079    ff1528114000            call dword ptr [00401128]
    Err.Raise 9
'0072307f    8b4dd4                  mov ecx, dword ptr [ebp-2c]
    
End If
'00723082    8d043f                  lea eax, dword ptr [edi+edi]
'00723085    eb09                    jmp 723090

' *** Reference to "__vbaGenerateBoundsError"
'00723087    ff1528114000            call dword ptr [00401128]
Err.Raise 9
'0072308d    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00723090    8b490c                  mov ecx, dword ptr [ecx+0c]
'00723093    8b95d8feffff            mov edx, dword ptr [ebp+fffffed8]
'00723099    668b140a                mov dx, word ptr [ecx+edx]
'0072309d    6683c201                add dx, 01
var_num4 = @[(var_44[((-27~))]] + 1
'007230a1    0f80490a0000            jo 723af0
'007230a7    66891408                mov word ptr [ecx+eax], dx
'007230ab    668b4598                mov ax, word ptr [ebp-68]
'007230af    66050100                add ax, 0001
var_num1 = __vbaStrCopy + 1
'007230b3    0f80370a0000            jo 723af0
'007230b9    c745b401000000          mov dword ptr [ebp-4c], 00000001
'007230c0    894598                  mov dword ptr [ebp-68], eax
'007230c3    66837dac01              cmp word ptr [ebp-54], 01
'007230c8    0f8ee6010000            jle 7232b4

Do While (var_135 > 1)
'007230ce    668b4db4                mov cx, word ptr [ebp-4c]
'007230d2    668b55ac                mov dx, word ptr [ebp-54]
'007230d6    666bc902                imul cx, 02
    var_num3 = var_num1 * 2
'007230da    0f80100a0000            jo 723af0
'007230e0    6683c201                add dx, 01
    var_num4 = var_135 + 1
'007230e4    0f80060a0000            jo 723af0
'007230ea    0fbfc2                  movsx eax, dx
'007230ed    8985d4feffff            mov dword ptr [ebp+fffffed4], eax
'007230f3    db85d4feffff            fild dword ptr [ebp+fffffed4]
'007230f9    894db4                  mov dword ptr [ebp-4c], ecx
'007230fc    dd9dccfeffff            fstp qword ptr [ebp+fffffecc]
    'var_115 = (00)
'00723102    dd85ccfeffff            fld qword ptr [ebp+fffffecc]
'00723108    833d00b0720000          cmp dword ptr [0072b000], 00
'0072310f    7508                    jne 723119
    
    If (    vbNullChar = 0) Then
'00723111    dc35d8154000            fdiv qword ptr [004015d8]
'00723117    eb11                    jmp 72312a
    
End If
'00723119    ff35dc154000            push dword ptr [004015dc]
'0072311f    ff35d8154000            push dword ptr [004015d8]

' *** Reference to "_adj_fdiv_m64"
'00723125    e85a41ceff              call 407284
'0072312a    dfe0                    fnstsw ax
'0072312c    a80d                    test al, 0d
'0072312e    0f85b7090000            jne 723aeb

' *** Reference to "__vbaR8IntI2"
'00723134    ff15c0124000            call dword ptr [004012c0]
'0072313a    83fe11                  cmp esi, 11

' *** Reference to "__vbaGenerateBoundsError"
'0072313d    8b3d28114000            mov edi, dword ptr [00401128]
'00723143    8945ac                  mov dword ptr [ebp-54], eax
'00723146    7202                    jc 72314a

If (__vbaStrCopy >= 17) Then
'00723148    ffd7                    call edi
    Err.Raise 9
End If
'0072314a    83fe11                  cmp esi, 11
'0072314d    7217                    jc 723166

If (__vbaStrCopy >= 17) Then
'0072314f    ffd7                    call edi
    Err.Raise 9
'00723151    83fe11                  cmp esi, 11
'00723154    7210                    jc 723166
    
    If (    __vbaStrCopy >= 17) Then
'00723156    ffd7                    call edi
    Err.Raise 9
'00723158    83fe11                  cmp esi, 11
'0072315b    7209                    jc 723166
    
    If (    __vbaStrCopy >= 17) Then
'0072315d    ffd7                    call edi
    Err.Raise 9
'0072315f    83fe11                  cmp esi, 11
'00723162    7202                    jc 723166
    
    If (    __vbaStrCopy >= 17) Then
'00723164    ffd7                    call edi
    Err.Raise 9
End If
'00723166    8b4dc8                  mov ecx, dword ptr [ebp-38]
'00723169    8b14f1                  mov edx, dword ptr [esi*8+ecx]

' *** Reference to "__vbaStrCmp"
'0072316c    8b3d30114000            mov edi, dword ptr [00401130]
'00723172    52                      push edx
'00723173    6800764500              push 00457600
'00723178    ffd7                    call edi
'0072317a    8bd0                    mov edx, eax
'0072317c    8b45c8                  mov eax, dword ptr [ebp-38]
'0072317f    8b0cf0                  mov ecx, dword ptr [esi*8+eax]
'00723182    f7da                    neg edx
'00723184    1bd2                    sbb edx, edx
'00723186    51                      push ecx
'00723187    f7da                    neg edx
'00723189    6880334300              push 00433380
'0072318e    8995c8feffff            mov dword ptr [ebp+fffffec8], edx
'00723194    ffd7                    call edi
'00723196    8bbdc8feffff            mov edi, dword ptr [ebp+fffffec8]
'0072319c    8b55c8                  mov edx, dword ptr [ebp-38]
'0072319f    f7d8                    neg eax
'007231a1    1bc0                    sbb eax, eax
'007231a3    f7d8                    neg eax
'007231a5    23f8                    and edi, eax
var_num7 = (("Adepte") [#$#] ("Expert")) And (("Adepte") [#$#] ("Adepte"))
'007231a7    8b04f2                  mov eax, dword ptr [esi*8+edx]
'007231aa    f7df                    neg edi
'007231ac    50                      push eax
'007231ad    1bff                    sbb edi, edi
'007231af    68d4624500              push 004562d4
'007231b4    f7df                    neg edi

' *** Reference to "__vbaStrCmp"
'007231b6    ff1530114000            call dword ptr [00401130]
'007231bc    8b4dc8                  mov ecx, dword ptr [ebp-38]
'007231bf    8b14f1                  mov edx, dword ptr [esi*8+ecx]
'007231c2    f7d8                    neg eax
'007231c4    1bc0                    sbb eax, eax
'007231c6    f7d8                    neg eax
'007231c8    23f8                    and edi, eax
var_num7 = -(CBool((var_num7))) And (("Adepte") [#$#] ("Gens du peuple"))
'007231ca    f7df                    neg edi
'007231cc    52                      push edx
'007231cd    1bff                    sbb edi, edi
'007231cf    6814764500              push 00457614
'007231d4    f7df                    neg edi

' *** Reference to "__vbaStrCmp"
'007231d6    ff1530114000            call dword ptr [00401130]
'007231dc    f7d8                    neg eax
'007231de    1bc0                    sbb eax, eax
'007231e0    f7d8                    neg eax
'007231e2    23f8                    and edi, eax
var_num7 = -(CBool((var_num7))) And (("Adepte") [#$#] ("Homme d'armes"))
'007231e4    8b45c8                  mov eax, dword ptr [ebp-38]
'007231e7    8b0cf0                  mov ecx, dword ptr [esi*8+eax]
'007231ea    f7df                    neg edi
'007231ec    51                      push ecx
'007231ed    1bff                    sbb edi, edi
'007231ef    68b86f4500              push 00456fb8
'007231f4    f7df                    neg edi

' *** Reference to "__vbaStrCmp"
'007231f6    ff1530114000            call dword ptr [00401130]
'007231fc    f7d8                    neg eax
'007231fe    1bc0                    sbb eax, eax
'00723200    f7d8                    neg eax
'00723202    85f8                    test ax, di
'00723204    750b                    jne 723211
'00723206    66837dac01              cmp word ptr [ebp-54], 01
'0072320b    0f84a3000000            je 7232b4

Do While ((81#  (var_51)) [##] 1)
'00723211    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00723214    85c9                    test cx, cx
'00723216    742a                    je 723242
'00723218    66833901                cmp word ptr [ecx], 01
'0072321c    7524                    jne 723242
'0072321e    0fbf7dac                movsx edi, word ptr [ebp-54]
'00723222    8b5114                  mov edx, dword ptr [ecx+14]
'00723225    8b4110                  mov eax, dword ptr [ecx+10]
'00723228    2bfa                    sub edi, edx
    var_num7 = (81# [:#] (var_51)) - LBound(var_44)
'0072322a    3bf8                    cmp edi, eax
'0072322c    7209                    jc 723237
    
    If (    var_num7 >= (UBound(var_44) - LBound(var_44))) Then

' *** Reference to "__vbaGenerateBoundsError"
'0072322e    ff1528114000            call dword ptr [00401128]
    Err.Raise 9
'00723234    8b4dd4                  mov ecx, dword ptr [ebp-2c]
    
End If
'00723237    8d143f                  lea edx, dword ptr [edi+edi]
'0072323a    8995c4feffff            mov dword ptr [ebp+fffffec4], edx
'00723240    eb0f                    jmp 723251

' *** Reference to "__vbaGenerateBoundsError"
'00723242    ff1528114000            call dword ptr [00401128]
Err.Raise 9
'00723248    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'0072324b    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'00723251    85c9                    test cx, cx
'00723253    7424                    je 723279
'00723255    66833901                cmp word ptr [ecx], 01
'00723259    751e                    jne 723279
'0072325b    0fbf7dac                movsx edi, word ptr [ebp-54]
'0072325f    8b5114                  mov edx, dword ptr [ecx+14]
'00723262    8b4110                  mov eax, dword ptr [ecx+10]
'00723265    2bfa                    sub edi, edx
var_num7 = (81# [:#] (var_51)) - LBound(var_44)
'00723267    3bf8                    cmp edi, eax
'00723269    7209                    jc 723274

If (var_num7 >= (UBound(var_44) - LBound(var_44))) Then

' *** Reference to "__vbaGenerateBoundsError"
'0072326b    ff1528114000            call dword ptr [00401128]
    Err.Raise 9
'00723271    8b4dd4                  mov ecx, dword ptr [ebp-2c]
    
End If
'00723274    8d043f                  lea eax, dword ptr [edi+edi]
'00723277    eb09                    jmp 723282

' *** Reference to "__vbaGenerateBoundsError"
'00723279    ff1528114000            call dword ptr [00401128]
Err.Raise 9
'0072327f    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00723282    8b490c                  mov ecx, dword ptr [ecx+0c]
'00723285    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'0072328b    668b3c0a                mov di, word ptr [ecx+edx]
'0072328f    8b55b4                  mov edx, dword ptr [ebp-4c]
'00723292    6603fa                  add di, dx
var_num7 = @[(var_44[((-27~))]] + var_num3
'00723295    0f8055080000            jo 723af0
'0072329b    66893c08                mov word ptr [ecx+eax], di
'0072329f    668b4598                mov ax, word ptr [ebp-68]
'007232a3    6603c2                  add ax, dx
var_num1 = var_num1 + var_num3
'007232a6    0f8044080000            jo 723af0
'007232ac    894598                  mov dword ptr [ebp-68], eax
'007232af    e90ffeffff              jmp 7230c3

'ERROR: Two many next close:
Loop

' *** Reference to "__vbaStrCat"
'007232b4    8b3d70104000            mov edi, dword ptr [00401070]
'007232ba    b801000000              mov eax, 00000001
'007232bf    660345e8                add ax, word ptr [ebp-18]
var_num1 = 1 + var_41
'007232c3    0f8027080000            jo 723af0
'007232c9    8945e8                  mov dword ptr [ebp-18], eax
'007232cc    e9c4fcffff              jmp 722f95

'ERROR: Two many next close:
Loop
'007232d1    66837dac00              cmp word ptr [ebp-54], 00
'007232d6    c745dc00000000          mov dword ptr [ebp-24], 00000000
'007232dd    0f8e69010000            jle 72344c

Do While ((81# [:#] (var_51)) > 0)
'007232e3    0fbf75e0                movsx esi, word ptr [ebp-20]
'007232e7    83fe11                  cmp esi, 11
'007232ea    7206                    jc 7232f2
    
    If (    __vbaStrCopy >= 17) Then

' *** Reference to "__vbaGenerateBoundsError"
'007232ec    ff1528114000            call dword ptr [00401128]
    Err.Raise 9
End If
'007232f2    8b4da8                  mov ecx, dword ptr [ebp-58]
'007232f5    8b55c8                  mov edx, dword ptr [ebp-38]
'007232f8    8b04f2                  mov eax, dword ptr [esi*8+edx]
'007232fb    51                      push ecx
'007232fc    50                      push eax
'007232fd    ffd7                    call edi
var_2587 = (vbNullChar) & ("Adepte")
'007232ff    8bd0                    mov edx, eax
'00723301    8d4d90                  lea ecx, dword ptr [ebp-70]
'00723304    ffd3                    call ebx
'00723306    50                      push eax
'00723307    68e42b4400              push 00442be4
'0072330c    ffd7                    call edi
var_53 = (var_2587) & (" : ")
'0072330e    8bd0                    mov edx, eax
'00723310    8d4da8                  lea ecx, dword ptr [ebp-58]
'00723313    ffd3                    call ebx
'00723315    8d4d90                  lea ecx, dword ptr [ebp-70]

' *** Reference to "__vbaFreeStr"
'00723318    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_8)
'0072331e    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00723321    c745e81c000000          mov dword ptr [ebp-18], 0000001c
'00723328    b801000000              mov eax, 00000001
'0072332d    663945e8                cmp word ptr [ebp-18], ax
'00723331    0f8c15010000            jl 72344c

Do While (var_41 >= 1)
'00723337    85c9                    test cx, cx
'00723339    7423                    je 72335e
'0072333b    663901                  cmp word ptr [ecx], ax
'0072333e    751e                    jne 72335e
'00723340    0fbf75e8                movsx esi, word ptr [ebp-18]
'00723344    8b5114                  mov edx, dword ptr [ecx+14]
'00723347    8b4110                  mov eax, dword ptr [ecx+10]
'0072334a    2bf2                    sub esi, edx
    var_num8 = var_41 - LBound(var_44)
'0072334c    3bf0                    cmp esi, eax
'0072334e    7209                    jc 723359
    
    If (    var_num8 >= (UBound(var_44) - LBound(var_44))) Then

' *** Reference to "__vbaGenerateBoundsError"
'00723350    ff1528114000            call dword ptr [00401128]
    Err.Raise 9
'00723356    8b4dd4                  mov ecx, dword ptr [ebp-2c]
    
End If
'00723359    8d0436                  lea eax, dword ptr [esi+esi]
'0072335c    eb09                    jmp 723367

' *** Reference to "__vbaGenerateBoundsError"
'0072335e    ff1528114000            call dword ptr [00401128]
Err.Raise 9
'00723364    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00723367    8b510c                  mov edx, dword ptr [ecx+0c]
'0072336a    66833c0200              cmp word ptr [eax+edx], 00
'0072336f    0f84c2000000            je 723437
'00723375    66837ddc00              cmp word ptr [ebp-24], 00
'0072337a    7417                    je 723393

If (var_39 <> 0) Then
'0072337c    8b45a8                  mov eax, dword ptr [ebp-58]
'0072337f    50                      push eax
'00723380    68fc8c4300              push 00438cfc
'00723385    ffd7                    call edi
    var_285 = (var_53) & (", ")
'00723387    8bd0                    mov edx, eax
'00723389    8d4da8                  lea ecx, dword ptr [ebp-58]
'0072338c    ffd3                    call ebx
'0072338e    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00723391    eb07                    jmp 72339a
    
End If
'00723393    c745dc01000000          mov dword ptr [ebp-24], 00000001
'0072339a    85c9                    test cx, cx
'0072339c    7424                    je 7233c2
'0072339e    66833901                cmp word ptr [ecx], 01
'007233a2    751e                    jne 7233c2
'007233a4    0fbf75e8                movsx esi, word ptr [ebp-18]
'007233a8    8b5114                  mov edx, dword ptr [ecx+14]
'007233ab    8b4110                  mov eax, dword ptr [ecx+10]
'007233ae    2bf2                    sub esi, edx
var_num8 = var_41 - LBound(var_44)
'007233b0    3bf0                    cmp esi, eax
'007233b2    7209                    jc 7233bd

If (var_num8 >= (UBound(var_44) - LBound(var_44))) Then

' *** Reference to "__vbaGenerateBoundsError"
'007233b4    ff1528114000            call dword ptr [00401128]
    Err.Raise 9
'007233ba    8b4dd4                  mov ecx, dword ptr [ebp-2c]
    
End If
'007233bd    8d0436                  lea eax, dword ptr [esi+esi]
'007233c0    eb09                    jmp 7233cb

' *** Reference to "__vbaGenerateBoundsError"
'007233c2    ff1528114000            call dword ptr [00401128]
Err.Raise 9
'007233c8    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'007233cb    8b55a8                  mov edx, dword ptr [ebp-58]
'007233ce    8b490c                  mov ecx, dword ptr [ecx+0c]

' *** Reference to "__vbaStrI2"
'007233d1    8b3510104000            mov esi, dword ptr [00401010]
'007233d7    52                      push edx
'007233d8    33d2                    xor edx, edx
var_num4 = Empty
'007233da    668b1401                mov dx, word ptr [eax+ecx]
'007233de    52                      push edx
'007233df    ffd6                    call esi
'007233e1    8bd0                    mov edx, eax
'007233e3    8d4d90                  lea ecx, dword ptr [ebp-70]
'007233e6    ffd3                    call ebx
'007233e8    50                      push eax
'007233e9    ffd7                    call edi
var_55 = (var_285) & (CStr(0))
'007233eb    8bd0                    mov edx, eax
'007233ed    8d4d8c                  lea ecx, dword ptr [ebp-74]
'007233f0    ffd3                    call ebx
'007233f2    50                      push eax
'007233f3    68bc484500              push 004548bc
'007233f8    ffd7                    call edi
var_139 = (var_55) & (" de niveau ")
'007233fa    8bd0                    mov edx, eax
'007233fc    8d4d88                  lea ecx, dword ptr [ebp-78]
'007233ff    ffd3                    call ebx
'00723401    50                      push eax
'00723402    8b45e8                  mov eax, dword ptr [ebp-18]
'00723405    50                      push eax
'00723406    ffd6                    call esi
'00723408    8bd0                    mov edx, eax
'0072340a    8d4d84                  lea ecx, dword ptr [ebp-7c]
'0072340d    ffd3                    call ebx
'0072340f    50                      push eax
'00723410    ffd7                    call edi
var_2462 = (var_139) & (CStr(var_41))
'00723412    8bd0                    mov edx, eax
'00723414    8d4da8                  lea ecx, dword ptr [ebp-58]
'00723417    ffd3                    call ebx
'00723419    8d4d84                  lea ecx, dword ptr [ebp-7c]
'0072341c    51                      push ecx
'0072341d    8d5588                  lea edx, dword ptr [ebp-78]
'00723420    52                      push edx
'00723421    8d458c                  lea eax, dword ptr [ebp-74]
'00723424    50                      push eax
'00723425    8d4d90                  lea ecx, dword ptr [ebp-70]
'00723428    51                      push ecx
'00723429    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'0072342b    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4560, -4564, -4568, -4572)
'00723431    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00723434    83c414                  add esp, 14
'00723437    83c8ff                  or eax, ffffffff
'0072343a    660345e8                add ax, word ptr [ebp-18]
var_num1 =  + var_41
'0072343e    0f80ac060000            jo 723af0
'00723444    8945e8                  mov dword ptr [ebp-18], eax
'00723447    e9dcfeffff              jmp 723328

'ERROR: Two many next close:
Loop
'0072344c    0fbf55e0                movsx edx, word ptr [ebp-20]
'00723450    8995c0feffff            mov dword ptr [ebp+fffffec0], edx
'00723456    c745dcffffffff          mov dword ptr [ebp-24], ffffffff
'0072345d    db85c0feffff            fild dword ptr [ebp+fffffec0]
'00723463    dd9db8feffff            fstp qword ptr [ebp+fffffeb8]
'var_767 = (00)
'00723469    dd459c                  fld qword ptr [ebp-64]

' *** Reference to "__vbaFPInt"
'0072346c    ff15fc124000            call dword ptr [004012fc]

' *** Reference to "__vbaFpR8"
'00723472    ff15f8104000            call dword ptr [004010f8]
'00723478    dc9db8feffff            fcomp qword ptr [ebp+fffffeb8]
'0072347e    dfe0                    fnstsw ax
'00723480    f6c440                  test ah, 40
'00723483    0f84b6000000            je 72353f

If (((__vbaStrCopy) = Int((var_51)))) Then
'00723489    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'0072348c    c745e81c000000          mov dword ptr [ebp-18], 0000001c
'00723493    b801000000              mov eax, 00000001
'00723498    663945e8                cmp word ptr [ebp-18], ax
'0072349c    0f8c9d000000            jl 72353f
    
    If (    var_41 >= 1) Then
'007234a2    85c9                    test cx, cx
'007234a4    7423                    je 7234c9
'007234a6    663901                  cmp word ptr [ecx], ax
'007234a9    751e                    jne 7234c9
'007234ab    0fbf75e8                movsx esi, word ptr [ebp-18]
'007234af    8b5114                  mov edx, dword ptr [ecx+14]
'007234b2    8b4110                  mov eax, dword ptr [ecx+10]
'007234b5    2bf2                    sub esi, edx
    var_num8 = var_41 - LBound(var_44)
'007234b7    3bf0                    cmp esi, eax
'007234b9    7209                    jc 7234c4
    
    If (    var_num8 >= (UBound(var_44) - LBound(var_44))) Then

' *** Reference to "__vbaGenerateBoundsError"
'007234bb    ff1528114000            call dword ptr [00401128]
    Err.Raise 9
'007234c1    8b4dd4                  mov ecx, dword ptr [ebp-2c]
    
End If
'007234c4    8d0436                  lea eax, dword ptr [esi+esi]
'007234c7    eb09                    jmp 7234d2

' *** Reference to "__vbaGenerateBoundsError"
'007234c9    ff1528114000            call dword ptr [00401128]
Err.Raise 9
'007234cf    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'007234d2    8b510c                  mov edx, dword ptr [ecx+0c]
'007234d5    66833c0200              cmp word ptr [eax+edx], 00
'007234da    7440                    je 72351c
'007234dc    dd459c                  fld qword ptr [ebp-64]

' *** Reference to "__vbaFPInt"
'007234df    ff15fc124000            call dword ptr [004012fc]
'007234e5    dc6d9c                  fsubr qword ptr [ebp-64]
'007234e8    dc0dd0194000            fmul qword ptr [004019d0]
'007234ee    dfe0                    fnstsw ax
'007234f0    a80d                    test al, 0d
'007234f2    0f85f3050000            jne 723aeb

' *** Reference to "__vbaFpR8"
'007234f8    ff15f8104000            call dword ptr [004010f8]
'007234fe    dc1d58164000            fcomp qword ptr [00401658]
'00723504    dfe0                    fnstsw ax
'00723506    f6c401                  test ah, 01
'00723509    7526                    jne 723531

Do While ((1# <= ((Int((var_51)) - var_51) * 10#)))
'0072350b    66837ddcff              cmp word ptr [ebp-24], ffffffff
'00723510    7527                    jne 723539
    
    If (    var_39 = -1) Then
'00723512    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'00723515    c745dc00000000          mov dword ptr [ebp-24], 00000000
'0072351c    83c8ff                  or eax, ffffffff
    var_num1 = (1# [:#] ((Int((var_51)) - var_51) * 10#)) Or -1
'0072351f    660345e8                add ax, word ptr [ebp-18]
    var_num1 = var_num1 + var_41
'00723523    0f80c7050000            jo 723af0
'00723529    8945e8                  mov dword ptr [ebp-18], eax
'0072352c    e962ffffff              jmp 723493
    
Loop
'00723531    8b45e8                  mov eax, dword ptr [ebp-18]
'00723534    8945a4                  mov dword ptr [ebp-5c], eax
'00723537    eb06                    jmp 72353f

'ERROR: Two many next close:
End If
'00723539    8b4de8                  mov ecx, dword ptr [ebp-18]
'0072353c    894da4                  mov dword ptr [ebp-5c], ecx

'ERROR: Two many next close:
End If
'0072353f    66837dac00              cmp word ptr [ebp-54], 00
'00723544    7e12                    jle 723558

If ((81# [:#] (var_51)) > 0) Then
'00723546    8b55a8                  mov edx, dword ptr [ebp-58]
'00723549    52                      push edx
'0072354a    6870084300              push 00430870
'0072354f    ffd7                    call edi
    var_137 = (var_2462) & (vbCrLf)
'00723551    8bd0                    mov edx, eax
'00723553    8d4da8                  lea ecx, dword ptr [ebp-58]
'00723556    ffd3                    call ebx
    
End If
'00723558    8b7508                  mov esi, dword ptr [ebp+08]
'0072355b    b801000000              mov eax, 00000001
'00723560    660345e0                add ax, word ptr [ebp-20]
var_num1 = 1 + __vbaStrCopy
'00723564    0f8086050000            jo 723af0
'0072356a    8945e0                  mov dword ptr [ebp-20], eax
'0072356d    e96ef9ffff              jmp 722ee0

'ERROR: Two many next close:
Loop
'00723572    8b06                    mov eax, dword ptr [esi]
'00723574    56                      push esi
'00723575    ff9018030000            call dword ptr [eax+00000318]
'0072357b    50                      push eax
'0072357c    8d4d80                  lea ecx, dword ptr [ebp-80]
'0072357f    51                      push ecx

' *** Reference to "__vbaObjSet"
'00723580    ff15b4104000            call dword ptr [004010b4]
Set var_18 = 1
'00723586    8bf0                    mov esi, eax
'00723588    8b16                    mov edx, dword ptr [esi]
'0072358a    8d4590                  lea eax, dword ptr [ebp-70]
'0072358d    50                      push eax
'0072358e    56                      push esi
'0072358f    ff92a0000000            call dword ptr [edx+000000a0]
'00723595    dbe2                    fnclex
'00723597    85c0                    test ax, ax
'00723599    7d12                    jge 7235ad
'0072359b    68a0000000              push 000000a0
'007235a0    68d00d4300              push 00430dd0
'007235a5    56                      push esi
'007235a6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007235a7    ff1580104000            call dword ptr [00401080]
'007235ad    8b4d90                  mov ecx, dword ptr [ebp-70]
'007235b0    51                      push ecx

' *** Reference to "__vbaR8Str"
'007235b1    ff1524124000            call dword ptr [00401224]
'007235b7    0fbf5598                movsx edx, word ptr [ebp-68]
'007235bb    8995b4feffff            mov dword ptr [ebp+fffffeb4], edx
'007235c1    db85b4feffff            fild dword ptr [ebp+fffffeb4]
'007235c7    dd9dacfeffff            fstp qword ptr [ebp+fffffeac]
'var_520 = (00)
'007235cd    dca5acfeffff            fsub qword ptr [ebp+fffffeac]
'007235d3    dfe0                    fnstsw ax
'007235d5    a80d                    test al, 0d
'007235d7    0f850e050000            jne 723aeb

' *** Reference to "__vbaFpI4"
'007235dd    ff15a4124000            call dword ptr [004012a4]
'var_num1 = CLng((CDbl(CStr(0)) - (var_num1)))
'007235e3    8d4d90                  lea ecx, dword ptr [ebp-70]
'007235e6    894594                  mov dword ptr [ebp-6c], eax

' *** Reference to "__vbaFreeStr"
'007235e9    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_8)
'007235ef    8d4d80                  lea ecx, dword ptr [ebp-80]

' *** Reference to "__vbaFreeObj"
'007235f2    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_18)
'007235f8    8b45a8                  mov eax, dword ptr [ebp-58]
'007235fb    50                      push eax
'007235fc    6870084300              push 00430870
'00723601    ffd7                    call edi
var_1231 = (var_137) & (vbCrLf)
'00723603    8bd0                    mov edx, eax
'00723605    8d4d90                  lea ecx, dword ptr [ebp-70]
'00723608    ffd3                    call ebx
'0072360a    50                      push eax
'0072360b    6834764500              push 00457634
'00723610    ffd7                    call edi
var_2091 = (var_1231) & ("Les classes de Pnj de niveau 1")
'00723612    8bd0                    mov edx, eax
'00723614    8d4d8c                  lea ecx, dword ptr [ebp-74]
'00723617    ffd3                    call ebx
'00723619    50                      push eax
'0072361a    6870084300              push 00430870
'0072361f    ffd7                    call edi
var_2069 = (var_2091) & (vbCrLf)
'00723621    8bd0                    mov edx, eax
'00723623    8d4da8                  lea ecx, dword ptr [ebp-58]
'00723626    ffd3                    call ebx
'00723628    8d4d8c                  lea ecx, dword ptr [ebp-74]
'0072362b    51                      push ecx
'0072362c    8d5590                  lea edx, dword ptr [ebp-70]
'0072362f    52                      push edx
'00723630    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'00723632    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4588, -4592)
'00723638    8b45a8                  mov eax, dword ptr [ebp-58]
'0072363b    83c40c                  add esp, 0c
'0072363e    50                      push eax
'0072363f    6878764500              push 00457678
'00723644    ffd7                    call edi
var_2092 = (var_2069) & ("Gens du peuple : ")
'00723646    8bd0                    mov edx, eax
'00723648    8d4d90                  lea ecx, dword ptr [ebp-70]
'0072364b    ffd3                    call ebx
'0072364d    db4594                  fild dword ptr [ebp-6c]
'00723650    50                      push eax
'00723651    dd9da4feffff            fstp qword ptr [ebp+fffffea4]
'var_534 = (00)
'00723657    dd85a4feffff            fld qword ptr [ebp+fffffea4]
'0072365d    dc0dd0704000            fmul qword ptr [004070d0]
'00723663    dfe0                    fnstsw ax
'00723665    a80d                    test al, 0d
'00723667    0f857e040000            jne 723aeb

' *** Reference to "__vbaFPInt"
'0072366d    ff15fc124000            call dword ptr [004012fc]

' *** Reference to "__vbaStrR8"
'00723673    8b3580114000            mov esi, dword ptr [00401180]
'00723679    83ec08                  sub esp, 08
'0072367c    dd1c24                  fstp qword ptr [esp]
'var_73 = (00)
'0072367f    ffd6                    call esi
'00723681    8bd0                    mov edx, eax
'00723683    8d4d8c                  lea ecx, dword ptr [ebp-74]
'00723686    ffd3                    call ebx
'00723688    50                      push eax
'00723689    ffd7                    call edi
var_2461 = (var_2092) & (CStr(Int((((CLng((CDbl(CStr(0)) - (var_num1))))) * 0.91#))))
'0072368b    8bd0                    mov edx, eax
'0072368d    8d4d88                  lea ecx, dword ptr [ebp-78]
'00723690    ffd3                    call ebx
'00723692    50                      push eax
'00723693    6870084300              push 00430870
'00723698    ffd7                    call edi
var_870 = (var_2461) & (vbCrLf)
'0072369a    8bd0                    mov edx, eax
'0072369c    8d4da8                  lea ecx, dword ptr [ebp-58]
'0072369f    ffd3                    call ebx
'007236a1    8d4d88                  lea ecx, dword ptr [ebp-78]
'007236a4    51                      push ecx
'007236a5    8d558c                  lea edx, dword ptr [ebp-74]
'007236a8    52                      push edx
'007236a9    8d4590                  lea eax, dword ptr [ebp-70]
'007236ac    50                      push eax
'007236ad    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'007236af    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4600, -4604, -4608)
'007236b5    8b4da8                  mov ecx, dword ptr [ebp-58]
'007236b8    83c410                  add esp, 10
'007236bb    51                      push ecx
'007236bc    68a0764500              push 004576a0
'007236c1    ffd7                    call edi
var_8 = (var_870) & ("Homme d'armes : ")
'007236c3    8bd0                    mov edx, eax
'007236c5    8d4d90                  lea ecx, dword ptr [ebp-70]
'007236c8    ffd3                    call ebx
'007236ca    db4594                  fild dword ptr [ebp-6c]
'007236cd    50                      push eax
'007236ce    dd9d9cfeffff            fstp qword ptr [ebp+fffffe9c]
'var_822 = (00)
'007236d4    dd859cfeffff            fld qword ptr [ebp+fffffe9c]
'007236da    dc0d58704000            fmul qword ptr [00407058]
'007236e0    dfe0                    fnstsw ax
'007236e2    a80d                    test al, 0d
'007236e4    0f8501040000            jne 723aeb

' *** Reference to "__vbaFPInt"
'007236ea    ff15fc124000            call dword ptr [004012fc]
'007236f0    83ec08                  sub esp, 08
'007236f3    dd1c24                  fstp qword ptr [esp]
'var_281 = (00)
'007236f6    ffd6                    call esi
'007236f8    8bd0                    mov edx, eax
'007236fa    8d4d8c                  lea ecx, dword ptr [ebp-74]
'007236fd    ffd3                    call ebx
'007236ff    50                      push eax
'00723700    ffd7                    call edi
var_871 = (var_8) & (CStr(Int((((CLng((CDbl(CStr(0)) - (var_num1))))) * 0.05#))))
'00723702    8bd0                    mov edx, eax
'00723704    8d4d88                  lea ecx, dword ptr [ebp-78]
'00723707    ffd3                    call ebx
'00723709    50                      push eax
'0072370a    6870084300              push 00430870
'0072370f    ffd7                    call edi
var_79 = (var_871) & (vbCrLf)
'00723711    8bd0                    mov edx, eax
'00723713    8d4da8                  lea ecx, dword ptr [ebp-58]
'00723716    ffd3                    call ebx
'00723718    8d5588                  lea edx, dword ptr [ebp-78]
'0072371b    52                      push edx
'0072371c    8d458c                  lea eax, dword ptr [ebp-74]
'0072371f    50                      push eax
'00723720    8d4d90                  lea ecx, dword ptr [ebp-70]
'00723723    51                      push ecx
'00723724    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00723726    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4616, -4620, -4624)
'0072372c    8b55a8                  mov edx, dword ptr [ebp-58]
'0072372f    83c410                  add esp, 10
'00723732    52                      push edx
'00723733    68c8764500              push 004576c8
'00723738    ffd7                    call edi
var_52 = (var_79) & ("Expert : ")
'0072373a    8bd0                    mov edx, eax
'0072373c    8d4d90                  lea ecx, dword ptr [ebp-70]
'0072373f    ffd3                    call ebx
'00723741    db4594                  fild dword ptr [ebp-6c]
'00723744    50                      push eax
'00723745    dd9d94feffff            fstp qword ptr [ebp+fffffe94]
'var_538 = (00)
'0072374b    dd8594feffff            fld qword ptr [ebp+fffffe94]
'00723751    dc0dc8704000            fmul qword ptr [004070c8]
'00723757    dfe0                    fnstsw ax
'00723759    a80d                    test al, 0d
'0072375b    0f858a030000            jne 723aeb

' *** Reference to "__vbaFPInt"
'00723761    ff15fc124000            call dword ptr [004012fc]
'00723767    83ec08                  sub esp, 08
'0072376a    dd1c24                  fstp qword ptr [esp]
'var_537 = (00)
'0072376d    ffd6                    call esi
'0072376f    8bd0                    mov edx, eax
'00723771    8d4d8c                  lea ecx, dword ptr [ebp-74]
'00723774    ffd3                    call ebx
'00723776    50                      push eax
'00723777    ffd7                    call edi
var_67 = (var_52) & (CStr(Int((((CLng((CDbl(CStr(0)) - (var_num1))))) * 0.03#))))
'00723779    8bd0                    mov edx, eax
'0072377b    8d4d88                  lea ecx, dword ptr [ebp-78]
'0072377e    ffd3                    call ebx
'00723780    50                      push eax
'00723781    6870084300              push 00430870
'00723786    ffd7                    call edi
var_68 = (var_67) & (vbCrLf)
'00723788    8bd0                    mov edx, eax
'0072378a    8d4da8                  lea ecx, dword ptr [ebp-58]
'0072378d    ffd3                    call ebx
'0072378f    8d4588                  lea eax, dword ptr [ebp-78]
'00723792    50                      push eax
'00723793    8d4d8c                  lea ecx, dword ptr [ebp-74]
'00723796    51                      push ecx
'00723797    8d5590                  lea edx, dword ptr [ebp-70]
'0072379a    52                      push edx
'0072379b    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'0072379d    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4632, -4636, -4640)
'007237a3    8b45a8                  mov eax, dword ptr [ebp-58]
'007237a6    83c410                  add esp, 10
'007237a9    50                      push eax
'007237aa    68e0764500              push 004576e0
'007237af    ffd7                    call edi
var_2062 = (var_68) & ("Adepte : ")
'007237b1    8bd0                    mov edx, eax
'007237b3    8d4d90                  lea ecx, dword ptr [ebp-70]
'007237b6    ffd3                    call ebx
'007237b8    db4594                  fild dword ptr [ebp-6c]
'007237bb    50                      push eax
'007237bc    dd9d8cfeffff            fstp qword ptr [ebp+fffffe8c]
'var_772 = (00)
'007237c2    dd858cfeffff            fld qword ptr [ebp+fffffe8c]
'007237c8    dc0dc0704000            fmul qword ptr [004070c0]
'007237ce    dfe0                    fnstsw ax
'007237d0    a80d                    test al, 0d
'007237d2    0f8513030000            jne 723aeb

' *** Reference to "__vbaFPInt"
'007237d8    ff15fc124000            call dword ptr [004012fc]
'007237de    83ec08                  sub esp, 08
'007237e1    dd1c24                  fstp qword ptr [esp]
'var_519 = (00)
'007237e4    ffd6                    call esi
'007237e6    8bd0                    mov edx, eax
'007237e8    8d4d8c                  lea ecx, dword ptr [ebp-74]
'007237eb    ffd3                    call ebx
'007237ed    50                      push eax
'007237ee    ffd7                    call edi
var_70 = (var_2062) & (CStr(Int((((CLng((CDbl(CStr(0)) - (var_num1))))) * 0.005#))))
'007237f0    8bd0                    mov edx, eax
'007237f2    8d4d88                  lea ecx, dword ptr [ebp-78]
'007237f5    ffd3                    call ebx
'007237f7    50                      push eax
'007237f8    6870084300              push 00430870
'007237fd    ffd7                    call edi
var_872 = (var_70) & (vbCrLf)
'007237ff    8bd0                    mov edx, eax
'00723801    8d4da8                  lea ecx, dword ptr [ebp-58]
'00723804    ffd3                    call ebx
'00723806    8d4d88                  lea ecx, dword ptr [ebp-78]
'00723809    51                      push ecx
'0072380a    8d558c                  lea edx, dword ptr [ebp-74]
'0072380d    52                      push edx
'0072380e    8d4590                  lea eax, dword ptr [ebp-70]
'00723811    50                      push eax
'00723812    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00723814    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4648, -4652, -4656)
'0072381a    8b4da8                  mov ecx, dword ptr [ebp-58]
'0072381d    83c410                  add esp, 10
'00723820    51                      push ecx
'00723821    68f8764500              push 004576f8
'00723826    ffd7                    call edi
var_8 = (var_872) & ("Noble : ")
'00723828    8bd0                    mov edx, eax
'0072382a    8d4d90                  lea ecx, dword ptr [ebp-70]
'0072382d    ffd3                    call ebx
'0072382f    db4594                  fild dword ptr [ebp-6c]
'00723832    50                      push eax
'00723833    dd9d84feffff            fstp qword ptr [ebp+fffffe84]
'var_282 = (00)
'00723839    dd8584feffff            fld qword ptr [ebp+fffffe84]
'0072383f    dc0dc0704000            fmul qword ptr [004070c0]
'00723845    dfe0                    fnstsw ax
'00723847    a80d                    test al, 0d
'00723849    0f859c020000            jne 723aeb

' *** Reference to "__vbaFPInt"
'0072384f    ff15fc124000            call dword ptr [004012fc]
'00723855    83ec08                  sub esp, 08
'00723858    dd1c24                  fstp qword ptr [esp]
'var_767 = (00)
'0072385b    ffd6                    call esi
'0072385d    8bd0                    mov edx, eax
'0072385f    8d4d8c                  lea ecx, dword ptr [ebp-74]
'00723862    ffd3                    call ebx
'00723864    50                      push eax
'00723865    ffd7                    call edi
var_2583 = (var_8) & (CStr(Int((((CLng((CDbl(CStr(0)) - (var_num1))))) * 0.005#))))
'00723867    8bd0                    mov edx, eax
'00723869    8d4d88                  lea ecx, dword ptr [ebp-78]
'0072386c    ffd3                    call ebx
'0072386e    50                      push eax
'0072386f    6870084300              push 00430870
'00723874    ffd7                    call edi
var_875 = (var_2583) & (vbCrLf)
'00723876    8bd0                    mov edx, eax
'00723878    8d4da8                  lea ecx, dword ptr [ebp-58]
'0072387b    ffd3                    call ebx
'0072387d    8d5588                  lea edx, dword ptr [ebp-78]
'00723880    52                      push edx
'00723881    8d458c                  lea eax, dword ptr [ebp-74]
'00723884    50                      push eax

' *** Reference to "__vbaFreeStrList"
'00723885    8b355c124000            mov esi, dword ptr [0040125c]
'0072388b    8d4d90                  lea ecx, dword ptr [ebp-70]
'0072388e    51                      push ecx
'0072388f    6a03                    push 03
'00723891    ffd6                    call esi
'#Cleanup( -4664, -4668, -4672)
'00723893    8b55a8                  mov edx, dword ptr [ebp-58]
'00723896    83c410                  add esp, 10
'00723899    52                      push edx
'0072389a    6870084300              push 00430870
'0072389f    ffd7                    call edi
var_52 = (var_875) & (vbCrLf)
'007238a1    8bd0                    mov edx, eax
'007238a3    8d4d90                  lea ecx, dword ptr [ebp-70]
'007238a6    ffd3                    call ebx
'007238a8    50                      push eax
'007238a9    6810774500              push 00457710
'007238ae    ffd7                    call edi
var_877 = (var_52) & ("Commandant de la garde/prévôt")
'007238b0    8bd0                    mov edx, eax
'007238b2    8d4d8c                  lea ecx, dword ptr [ebp-74]
'007238b5    ffd3                    call ebx
'007238b7    50                      push eax
'007238b8    6870084300              push 00430870
'007238bd    ffd7                    call edi
var_878 = (var_877) & (vbCrLf)
'007238bf    8bd0                    mov edx, eax
'007238c1    8d4da8                  lea ecx, dword ptr [ebp-58]
'007238c4    ffd3                    call ebx
'007238c6    8d458c                  lea eax, dword ptr [ebp-74]
'007238c9    50                      push eax
'007238ca    8d4d90                  lea ecx, dword ptr [ebp-70]
'007238cd    51                      push ecx
'007238ce    6a02                    push 02
'007238d0    ffd6                    call esi
'#Cleanup( -4680, -4684)
'007238d2    83c40c                  add esp, 0c
'007238d5    66837da400              cmp word ptr [ebp-5c], 00
'007238da    0f8e3f010000            jle 723a1f

If (var_num1 > 0) Then
'007238e0    8b4d9c                  mov ecx, dword ptr [ebp-64]
'007238e3    81f933333333            cmp ecx, 33333333
'007238e9    8b45a0                  mov eax, dword ptr [ebp-60]
'007238ec    755c                    jne 72394a
    
    If (    var_51 = 858993459) Then
'007238ee    3d33332040              cmp eax, 40203333
'007238f3    7555                    jne 72394a
    
    If (    var_7 = 1075852083) Then
'007238f5    8b55a8                  mov edx, dword ptr [ebp-58]
'007238f8    52                      push edx
'007238f9    6850774500              push 00457750
'007238fe    ffd7                    call edi
    var_2588 = (var_878) & ("Homme d'armes du plus haut niveau ")
'00723900    8bd0                    mov edx, eax
'00723902    8d4d90                  lea ecx, dword ptr [ebp-70]
'00723905    ffd3                    call ebx
'00723907    50                      push eax
'00723908    8b45a4                  mov eax, dword ptr [ebp-5c]
'0072390b    50                      push eax

' *** Reference to "__vbaStrI2"
'0072390c    ff1510104000            call dword ptr [00401010]
'00723912    8bd0                    mov edx, eax
'00723914    8d4d8c                  lea ecx, dword ptr [ebp-74]
'00723917    ffd3                    call ebx
'00723919    50                      push eax
'0072391a    ffd7                    call edi
    var_880 = (var_2588) & (CStr(var_num1))
'0072391c    8bd0                    mov edx, eax
'0072391e    8d4d88                  lea ecx, dword ptr [ebp-78]
'00723921    ffd3                    call ebx
'00723923    50                      push eax
'00723924    6870084300              push 00430870
'00723929    ffd7                    call edi
    var_1386 = (var_880) & (vbCrLf)
'0072392b    8bd0                    mov edx, eax
'0072392d    8d4da8                  lea ecx, dword ptr [ebp-58]
'00723930    ffd3                    call ebx
'00723932    8d4d88                  lea ecx, dword ptr [ebp-78]
'00723935    51                      push ecx
'00723936    8d558c                  lea edx, dword ptr [ebp-74]
'00723939    52                      push edx
'0072393a    8d4590                  lea eax, dword ptr [ebp-70]
'0072393d    50                      push eax
'0072393e    6a03                    push 03
'00723940    ffd6                    call esi
    '#Cleanup( -4692, -4696, -4700)
'00723942    83c410                  add esp, 10
'00723945    e9ff000000              jmp 723a49
    
End If
'0072394a    81f9cdcccccc            cmp ecx, cccccccd
'00723950    7560                    jne 7239b2

If (var_51 = -858993459) Then
'00723952    3dcccc1c40              cmp eax, 401ccccc
'00723957    7559                    jne 7239b2
    
    If (    var_7 = 1075629260) Then
'00723959    8b4da8                  mov ecx, dword ptr [ebp-58]
'0072395c    51                      push ecx
'0072395d    689c774500              push 0045779c
'00723962    ffd7                    call edi
    var_2588 = (var_1386) & ("Deuxième guerrier de la communauté ")
'00723964    8bd0                    mov edx, eax
'00723966    8d4d90                  lea ecx, dword ptr [ebp-70]
'00723969    ffd3                    call ebx
'0072396b    8b55a4                  mov edx, dword ptr [ebp-5c]
'0072396e    50                      push eax
'0072396f    52                      push edx

' *** Reference to "__vbaStrI2"
'00723970    ff1510104000            call dword ptr [00401010]
'00723976    8bd0                    mov edx, eax
'00723978    8d4d8c                  lea ecx, dword ptr [ebp-74]
'0072397b    ffd3                    call ebx
'0072397d    50                      push eax
'0072397e    ffd7                    call edi
    var_296 = (var_2588) & (CStr(var_num1))
'00723980    8bd0                    mov edx, eax
'00723982    8d4d88                  lea ecx, dword ptr [ebp-78]
'00723985    ffd3                    call ebx
'00723987    50                      push eax
'00723988    6870084300              push 00430870
'0072398d    ffd7                    call edi
    var_883 = (var_296) & (vbCrLf)
'0072398f    8bd0                    mov edx, eax
'00723991    8d4da8                  lea ecx, dword ptr [ebp-58]
'00723994    ffd3                    call ebx
'00723996    8d4588                  lea eax, dword ptr [ebp-78]
'00723999    50                      push eax
'0072399a    8d4d8c                  lea ecx, dword ptr [ebp-74]
'0072399d    51                      push ecx
'0072399e    8d5590                  lea edx, dword ptr [ebp-70]
'007239a1    52                      push edx
'007239a2    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'007239a4    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4708, -4712, -4716)
'007239aa    83c410                  add esp, 10
'007239ad    e997000000              jmp 723a49
    
End If
'007239b2    81f966666666            cmp ecx, 66666666
'007239b8    0f858b000000            jne 723a49

If (var_51 = 1717986918) Then
'007239be    3d66661c40              cmp eax, 401c6666
'007239c3    0f8580000000            jne 723a49
    
    If (    var_7 = 1075603046) Then
'007239c9    8b45a8                  mov eax, dword ptr [ebp-58]
'007239cc    50                      push eax
'007239cd    6800784500              push 00457800
'007239d2    ffd7                    call edi
    var_884 = (var_883) & ("Guerrier de plus haut niveau ")
'007239d4    8bd0                    mov edx, eax
'007239d6    8d4d90                  lea ecx, dword ptr [ebp-70]
'007239d9    ffd3                    call ebx
'007239db    8b4da4                  mov ecx, dword ptr [ebp-5c]
'007239de    50                      push eax
'007239df    51                      push ecx

' *** Reference to "__vbaStrI2"
'007239e0    ff1510104000            call dword ptr [00401010]
'007239e6    8bd0                    mov edx, eax
'007239e8    8d4d8c                  lea ecx, dword ptr [ebp-74]
'007239eb    ffd3                    call ebx
'007239ed    50                      push eax
'007239ee    ffd7                    call edi
    var_886 = (var_884) & (CStr(var_num1))
'007239f0    8bd0                    mov edx, eax
'007239f2    8d4d88                  lea ecx, dword ptr [ebp-78]
'007239f5    ffd3                    call ebx
'007239f7    50                      push eax
'007239f8    6870084300              push 00430870
'007239fd    ffd7                    call edi
    var_885 = (var_886) & (vbCrLf)
'007239ff    8bd0                    mov edx, eax
'00723a01    8d4da8                  lea ecx, dword ptr [ebp-58]
'00723a04    ffd3                    call ebx
'00723a06    8d5588                  lea edx, dword ptr [ebp-78]
'00723a09    52                      push edx
'00723a0a    8d458c                  lea eax, dword ptr [ebp-74]
'00723a0d    50                      push eax
'00723a0e    8d4d90                  lea ecx, dword ptr [ebp-70]
'00723a11    51                      push ecx
'00723a12    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00723a14    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4724, -4728, -4732)
'00723a1a    83c410                  add esp, 10
'00723a1d    eb2a                    jmp 723a49
    
End If
'00723a1f    8b55a8                  mov edx, dword ptr [ebp-58]
'00723a22    52                      push edx
'00723a23    6840784500              push 00457840
'00723a28    ffd7                    call edi
var_52 = (var_885) & ("Il n'y a aucun guerrier ou homme d'armes dans cette communauté, choisir un autre PNJ")
'00723a2a    8bd0                    mov edx, eax
'00723a2c    8d4d90                  lea ecx, dword ptr [ebp-70]
'00723a2f    ffd3                    call ebx
'00723a31    50                      push eax
'00723a32    6870084300              push 00430870
'00723a37    ffd7                    call edi
var_2369 = (var_52) & (vbCrLf)
'00723a39    8bd0                    mov edx, eax
'00723a3b    8d4da8                  lea ecx, dword ptr [ebp-58]
'00723a3e    ffd3                    call ebx
'00723a40    8d4d90                  lea ecx, dword ptr [ebp-70]

' *** Reference to "__vbaFreeStr"
'00723a43    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_8)
'ERROR: Two many next close:
End If
'00723a49    9b                      fwait
'00723a4a    68c43a7200              push 00723ac4
'00723a4f    eb4d                    jmp 723a9e
'00723a51    f645fc04                test byte ptr [ebp-04], 04
'00723a55    7409                    je 723a60
'00723a57    8d4da8                  lea ecx, dword ptr [ebp-58]

' *** Reference to "__vbaFreeStr"
'00723a5a    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_86)
'00723a60    8d4584                  lea eax, dword ptr [ebp-7c]
'00723a63    50                      push eax
'00723a64    8d4d88                  lea ecx, dword ptr [ebp-78]
'00723a67    51                      push ecx
'00723a68    8d558c                  lea edx, dword ptr [ebp-74]
'00723a6b    52                      push edx
'00723a6c    8d4590                  lea eax, dword ptr [ebp-70]
'00723a6f    50                      push eax
'00723a70    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'00723a72    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , -4728, -4732, -4572)
'00723a78    83c414                  add esp, 14
'00723a7b    8d4d80                  lea ecx, dword ptr [ebp-80]

' *** Reference to "__vbaFreeObj"
'00723a7e    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_18)
'00723a84    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'00723a8a    51                      push ecx
'00723a8b    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'00723a91    52                      push edx
'00723a92    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'00723a94    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_19, var_20)
'00723a9a    83c40c                  add esp, 0c
'00723a9d    c3                      ret

' *** Reference to "__vbaAryDestruct"
'00723a9e    8b3594104000            mov esi, dword ptr [00401094]
'00723aa4    8d45d4                  lea eax, dword ptr [ebp-2c]
'00723aa7    50                      push eax
'00723aa8    6a00                    push 00
'00723aaa    ffd6                    call esi
'00723aac    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'00723ab2    52                      push edx
'00723ab3    8d4dbc                  lea ecx, dword ptr [ebp-44]
'00723ab6    6898004300              push 00430098
'00723abb    898d2cffffff            mov dword ptr [ebp+ffffff2c], ecx
'00723ac1    ffd6                    call esi
'00723ac3    c3                      ret
'00723ac4    8b4508                  mov eax, dword ptr [ebp+08]
'00723ac7    8b08                    mov ecx, dword ptr [eax]
'00723ac9    50                      push eax
'00723aca    ff5108                  call dword ptr [ecx+08]
'00723acd    8b45a8                  mov eax, dword ptr [ebp-58]
'00723ad0    8b5510                  mov edx, dword ptr [ebp+10]
'00723ad3    8902                    mov dword ptr [edx], eax
'00723ad5    8b45fc                  mov eax, dword ptr [ebp-04]
'00723ad8    8b4dec                  mov ecx, dword ptr [ebp-14]
'00723adb    5f                      pop edi
'00723adc    5e                      pop esi
    'Reference to '__except_list'
'00723add    64890d00000000          mov dword ptr fs:[00000000], ecx
'00723ae4    5b                      pop ebx
'00723ae5    8be5                    mov esp, ebp
'00723ae7    5d                      pop ebp
'00723ae8    c20c00                  ret 000c


End Function


Public Function RACE(arg_0 As Boolean, arg_1 As Unknow, arg_2 As String, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'00723b00    55                      push ebp
'00723b01    8bec                    mov ebp, esp
'00723b03    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00723b06    6866724000              push 00407266
'00723b0b    64a100000000            mov ax, word ptr fs:[00000000]
'00723b11    50                      push eax
    'Reference to '__except_list'
'00723b12    64892500000000          mov dword ptr fs:[00000000], esp
'00723b19    81ecac000000            sub esp, 000000ac
'00723b1f    53                      push ebx
'00723b20    56                      push esi
'00723b21    57                      push edi
'00723b22    8965f4                  mov dword ptr [ebp-0c], esp
'00723b25    c745f840714000          mov dword ptr [ebp-08], 00407140
'00723b2c    33f6                    xor esi, esi
'00723b2e    8975fc                  mov dword ptr [ebp-04], esi
'00723b31    8b4508                  mov eax, dword ptr [ebp+08]
'00723b34    8b08                    mov ecx, dword ptr [eax]
'00723b36    50                      push eax
'00723b37    ff5104                  call dword ptr [ecx+04]
'00723b3a    8b5514                  mov edx, dword ptr [ebp+14]
'00723b3d    8932                    mov dword ptr [edx], esi
'00723b3f    8975e8                  mov dword ptr [ebp-18], esi
'00723b42    8975e4                  mov dword ptr [ebp-1c], esi
'00723b45    8975e0                  mov dword ptr [ebp-20], esi
'00723b48    8975dc                  mov dword ptr [ebp-24], esi
'00723b4b    8975d8                  mov dword ptr [ebp-28], esi

' *** Reference to "__vbaStrCopy"
'00723b4e    8b3548124000            mov esi, dword ptr [00401248]
'00723b54    bacc134300              mov edx, 004313cc
'00723b59    8d4de8                  lea ecx, dword ptr [ebp-18]
'00723b5c    ffd6                    call esi
var_41 = (vbNullChar)
'00723b5e    8b450c                  mov eax, dword ptr [ebp+0c]
'00723b61    8b10                    mov edx, dword ptr [eax]
'00723b63    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00723b66    ffd6                    call esi
var_45 = (arg_0)
'00723b68    8b4dd8                  mov ecx, dword ptr [ebp-28]

' *** Reference to "__vbaStrCmp"
'00723b6b    8b3530114000            mov esi, dword ptr [00401130]
'00723b71    51                      push ecx
'00723b72    680c794500              push 0045790c
'00723b77    ffd6                    call esi
'00723b79    85c0                    test ax, ax
'00723b7b    0f85d9010000            jne 723d5a

If (((arg_0) = ("Isolée"))) Then
'00723b81    8b55e8                  mov edx, dword ptr [ebp-18]

' *** Reference to "__vbaStrCat"
'00723b84    8b3d70104000            mov edi, dword ptr [00401070]
'00723b8a    52                      push edx
'00723b8b    6820794500              push 00457920
'00723b90    ffd7                    call edi
    var_84 = (vbNullChar) & ("Humains : ")

' *** Reference to "__vbaStrMove"
'00723b92    8b35d0124000            mov esi, dword ptr [004012d0]
'00723b98    8bd0                    mov edx, eax
'00723b9a    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00723b9d    ffd6                    call esi
'00723b9f    8b5d10                  mov ebx, dword ptr [ebp+10]
'00723ba2    db03                    fild dword ptr [ebx]
'00723ba4    50                      push eax
'00723ba5    dd5dc8                  fstp qword ptr [ebp-38]
    'var_46 = (00)
'00723ba8    dd45c8                  fld qword ptr [ebp-38]
'00723bab    dc0d38714000            fmul qword ptr [00407138]
'00723bb1    dfe0                    fnstsw ax
'00723bb3    a80d                    test al, 0d
'00723bb5    0f85c5080000            jne 724480

' *** Reference to "__vbaFPInt"
'00723bbb    ff15fc124000            call dword ptr [004012fc]
'00723bc1    83ec08                  sub esp, 08
'00723bc4    dd1c24                  fstp qword ptr [esp]
    'var_136 = (00)

' *** Reference to "__vbaStrR8"
'00723bc7    ff1580114000            call dword ptr [00401180]
'00723bcd    8bd0                    mov edx, eax
'00723bcf    8d4de0                  lea ecx, dword ptr [ebp-20]
'00723bd2    ffd6                    call esi
'00723bd4    50                      push eax
'00723bd5    ffd7                    call edi
    var_36 = (var_84) & (CStr(Int((((arg_1)) * 0.96#))))
'00723bd7    8bd0                    mov edx, eax
'00723bd9    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00723bdc    ffd6                    call esi
'00723bde    50                      push eax
'00723bdf    6870084300              push 00430870
'00723be4    ffd7                    call edi
    var_76 = (var_36) & (vbCrLf)
'00723be6    8bd0                    mov edx, eax
'00723be8    8d4de8                  lea ecx, dword ptr [ebp-18]
'00723beb    ffd6                    call esi
'00723bed    8d45dc                  lea eax, dword ptr [ebp-24]
'00723bf0    50                      push eax
'00723bf1    8d4de0                  lea ecx, dword ptr [ebp-20]
'00723bf4    51                      push ecx
'00723bf5    8d55e4                  lea edx, dword ptr [ebp-1c]
'00723bf8    52                      push edx
'00723bf9    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00723bfb    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4500, -4504, -4508)
'00723c01    8b45e8                  mov eax, dword ptr [ebp-18]
'00723c04    83c410                  add esp, 10
'00723c07    50                      push eax
'00723c08    683c794500              push 0045793c
'00723c0d    ffd7                    call edi
    var_12 = (var_76) & ("Halfelins : ")
'00723c0f    8bd0                    mov edx, eax
'00723c11    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00723c14    ffd6                    call esi
'00723c16    db03                    fild dword ptr [ebx]
'00723c18    50                      push eax
'00723c19    dd5dc0                  fstp qword ptr [ebp-40]
    'var_5 = (00)
'00723c1c    dd45c0                  fld qword ptr [ebp-40]
'00723c1f    dc0d30714000            fmul qword ptr [00407130]
'00723c25    dfe0                    fnstsw ax
'00723c27    a80d                    test al, 0d
'00723c29    0f8551080000            jne 724480

' *** Reference to "__vbaFPInt"
'00723c2f    ff15fc124000            call dword ptr [004012fc]
'00723c35    83ec08                  sub esp, 08
'00723c38    dd1c24                  fstp qword ptr [esp]
    'var_135 = (00)

' *** Reference to "__vbaStrR8"
'00723c3b    ff1580114000            call dword ptr [00401180]
'00723c41    8bd0                    mov edx, eax
'00723c43    8d4de0                  lea ecx, dword ptr [ebp-20]
'00723c46    ffd6                    call esi
'00723c48    50                      push eax
'00723c49    ffd7                    call edi
    var_14 = (var_12) & (CStr(Int((((arg_1)) * 0.02#))))
'00723c4b    8bd0                    mov edx, eax
'00723c4d    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00723c50    ffd6                    call esi
'00723c52    50                      push eax
'00723c53    6870084300              push 00430870
'00723c58    ffd7                    call edi
    var_127 = (var_14) & (vbCrLf)
'00723c5a    8bd0                    mov edx, eax
'00723c5c    8d4de8                  lea ecx, dword ptr [ebp-18]
'00723c5f    ffd6                    call esi
'00723c61    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00723c64    51                      push ecx
'00723c65    8d55e0                  lea edx, dword ptr [ebp-20]
'00723c68    52                      push edx
'00723c69    8d45e4                  lea eax, dword ptr [ebp-1c]
'00723c6c    50                      push eax
'00723c6d    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00723c6f    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4516, -4520, -4524)
'00723c75    8b4de8                  mov ecx, dword ptr [ebp-18]
'00723c78    83c410                  add esp, 10
'00723c7b    51                      push ecx
'00723c7c    685c794500              push 0045795c
'00723c81    ffd7                    call edi
    var_40 = (var_127) & ("Elfes : ")
'00723c83    8bd0                    mov edx, eax
'00723c85    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00723c88    ffd6                    call esi
'00723c8a    db03                    fild dword ptr [ebx]
'00723c8c    50                      push eax
'00723c8d    dd5db8                  fstp qword ptr [ebp-48]
    'var_61 = (00)
'00723c90    dd45b8                  fld qword ptr [ebp-48]
'00723c93    dc0d50704000            fmul qword ptr [00407050]
'00723c99    dfe0                    fnstsw ax
'00723c9b    a80d                    test al, 0d
'00723c9d    0f85dd070000            jne 724480

' *** Reference to "__vbaFPInt"
'00723ca3    ff15fc124000            call dword ptr [004012fc]
'00723ca9    83ec08                  sub esp, 08
'00723cac    dd1c24                  fstp qword ptr [esp]
    'var_134 = (00)

' *** Reference to "__vbaStrR8"
'00723caf    ff1580114000            call dword ptr [00401180]
'00723cb5    8bd0                    mov edx, eax
'00723cb7    8d4de0                  lea ecx, dword ptr [ebp-20]
'00723cba    ffd6                    call esi
'00723cbc    50                      push eax
'00723cbd    ffd7                    call edi
    var_128 = (var_40) & (CStr(Int((((arg_1)) * 0.1#))))
'00723cbf    8bd0                    mov edx, eax
'00723cc1    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00723cc4    ffd6                    call esi
'00723cc6    50                      push eax
'00723cc7    6870084300              push 00430870
'00723ccc    ffd7                    call edi
    var_17 = (var_128) & (vbCrLf)
'00723cce    8bd0                    mov edx, eax
'00723cd0    8d4de8                  lea ecx, dword ptr [ebp-18]
'00723cd3    ffd6                    call esi
'00723cd5    8d55dc                  lea edx, dword ptr [ebp-24]
'00723cd8    52                      push edx
'00723cd9    8d45e0                  lea eax, dword ptr [ebp-20]
'00723cdc    50                      push eax
'00723cdd    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00723ce0    51                      push ecx
'00723ce1    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00723ce3    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4532, -4536, -4540)
'00723ce9    8b55e8                  mov edx, dword ptr [ebp-18]
'00723cec    83c410                  add esp, 10
'00723cef    52                      push edx
'00723cf0    6874794500              push 00457974
'00723cf5    ffd7                    call edi
    var_3 = (var_17) & ("Autres races : ")
'00723cf7    8bd0                    mov edx, eax
'00723cf9    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00723cfc    ffd6                    call esi
'00723cfe    db03                    fild dword ptr [ebx]
'00723d00    50                      push eax
'00723d01    dd5db0                  fstp qword ptr [ebp-50]
    'var_6 = (00)
'00723d04    dd45b0                  fld qword ptr [ebp-50]
'00723d07    dc0d28714000            fmul qword ptr [00407128]
'00723d0d    dfe0                    fnstsw ax
'00723d0f    a80d                    test al, 0d
'00723d11    0f8569070000            jne 724480

' *** Reference to "__vbaFPInt"
'00723d17    ff15fc124000            call dword ptr [004012fc]
'00723d1d    83ec08                  sub esp, 08
'00723d20    dd1c24                  fstp qword ptr [esp]
    'var_133 = (00)

' *** Reference to "__vbaStrR8"
'00723d23    ff1580114000            call dword ptr [00401180]
'00723d29    8bd0                    mov edx, eax
'00723d2b    8d4de0                  lea ecx, dword ptr [ebp-20]
'00723d2e    ffd6                    call esi
'00723d30    50                      push eax
'00723d31    ffd7                    call edi
    var_285 = (var_3) & (CStr(Int((((arg_1)) * 0.01#))))
'00723d33    8bd0                    mov edx, eax
'00723d35    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00723d38    ffd6                    call esi
'00723d3a    50                      push eax
'00723d3b    6870084300              push 00430870
'00723d40    ffd7                    call edi
    var_54 = (var_285) & (vbCrLf)
'00723d42    8bd0                    mov edx, eax
'00723d44    8d4de8                  lea ecx, dword ptr [ebp-18]
'00723d47    ffd6                    call esi
'00723d49    8d45dc                  lea eax, dword ptr [ebp-24]
'00723d4c    50                      push eax
'00723d4d    8d4de0                  lea ecx, dword ptr [ebp-20]
'00723d50    51                      push ecx
'00723d51    8d55e4                  lea edx, dword ptr [ebp-1c]
'00723d54    52                      push edx
'00723d55    e9bb060000              jmp 724415
    
End If
'00723d5a    8b45d8                  mov eax, dword ptr [ebp-28]
'00723d5d    50                      push eax
'00723d5e    6898794500              push 00457998
'00723d63    ffd6                    call esi
'00723d65    85c0                    test ax, ax
'00723d67    0f853b030000            jne 7240a8

If (((arg_0) = ("Ouverte"))) Then
'00723d6d    8b4de8                  mov ecx, dword ptr [ebp-18]

' *** Reference to "__vbaStrCat"
'00723d70    8b3d70104000            mov edi, dword ptr [00401070]
'00723d76    51                      push ecx
'00723d77    6820794500              push 00457920
'00723d7c    ffd7                    call edi
    var_139 = (var_54) & ("Humains : ")

' *** Reference to "__vbaStrMove"
'00723d7e    8b35d0124000            mov esi, dword ptr [004012d0]
'00723d84    8bd0                    mov edx, eax
'00723d86    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00723d89    ffd6                    call esi
'00723d8b    8b5d10                  mov ebx, dword ptr [ebp+10]
'00723d8e    db03                    fild dword ptr [ebx]
'00723d90    50                      push eax
'00723d91    dd5da8                  fstp qword ptr [ebp-58]
    'var_86 = (00)
'00723d94    dd45a8                  fld qword ptr [ebp-58]
'00723d97    dc0d20714000            fmul qword ptr [00407120]
'00723d9d    dfe0                    fnstsw ax
'00723d9f    a80d                    test al, 0d
'00723da1    0f85d9060000            jne 724480

' *** Reference to "__vbaFPInt"
'00723da7    ff15fc124000            call dword ptr [004012fc]
'00723dad    83ec08                  sub esp, 08
'00723db0    dd1c24                  fstp qword ptr [esp]
    'var_136 = (00)

' *** Reference to "__vbaStrR8"
'00723db3    ff1580114000            call dword ptr [00401180]
'00723db9    8bd0                    mov edx, eax
'00723dbb    8d4de0                  lea ecx, dword ptr [ebp-20]
'00723dbe    ffd6                    call esi
'00723dc0    50                      push eax
'00723dc1    ffd7                    call edi
    var_2462 = (var_139) & (CStr(Int((((arg_1)) * 0.79#))))
'00723dc3    8bd0                    mov edx, eax
'00723dc5    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00723dc8    ffd6                    call esi
'00723dca    50                      push eax
'00723dcb    6870084300              push 00430870
'00723dd0    ffd7                    call edi
    var_286 = (var_2462) & (vbCrLf)
'00723dd2    8bd0                    mov edx, eax
'00723dd4    8d4de8                  lea ecx, dword ptr [ebp-18]
'00723dd7    ffd6                    call esi
'00723dd9    8d55dc                  lea edx, dword ptr [ebp-24]
'00723ddc    52                      push edx
'00723ddd    8d45e0                  lea eax, dword ptr [ebp-20]
'00723de0    50                      push eax
'00723de1    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00723de4    51                      push ecx
'00723de5    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00723de7    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4568, -4572, -4576)
'00723ded    8b55e8                  mov edx, dword ptr [ebp-18]
'00723df0    83c410                  add esp, 10
'00723df3    52                      push edx
'00723df4    683c794500              push 0045793c
'00723df9    ffd7                    call edi
    var_3 = (var_286) & ("Halfelins : ")
'00723dfb    8bd0                    mov edx, eax
'00723dfd    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00723e00    ffd6                    call esi
'00723e02    db03                    fild dword ptr [ebx]
'00723e04    50                      push eax
'00723e05    dd5da0                  fstp qword ptr [ebp-60]
    'var_7 = (00)
'00723e08    dd45a0                  fld qword ptr [ebp-60]
'00723e0b    dc0d18714000            fmul qword ptr [00407118]
'00723e11    dfe0                    fnstsw ax
'00723e13    a80d                    test al, 0d
'00723e15    0f8565060000            jne 724480

' *** Reference to "__vbaFPInt"
'00723e1b    ff15fc124000            call dword ptr [004012fc]
'00723e21    83ec08                  sub esp, 08
'00723e24    dd1c24                  fstp qword ptr [esp]
    'var_135 = (00)

' *** Reference to "__vbaStrR8"
'00723e27    ff1580114000            call dword ptr [00401180]
'00723e2d    8bd0                    mov edx, eax
'00723e2f    8d4de0                  lea ecx, dword ptr [ebp-20]
'00723e32    ffd6                    call esi
'00723e34    50                      push eax
'00723e35    ffd7                    call edi
    var_2091 = (var_3) & (CStr(Int((((arg_1)) * 0.09#))))
'00723e37    8bd0                    mov edx, eax
'00723e39    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00723e3c    ffd6                    call esi
'00723e3e    50                      push eax
'00723e3f    6870084300              push 00430870
'00723e44    ffd7                    call edi
    var_2069 = (var_2091) & (vbCrLf)
'00723e46    8bd0                    mov edx, eax
'00723e48    8d4de8                  lea ecx, dword ptr [ebp-18]
'00723e4b    ffd6                    call esi
'00723e4d    8d45dc                  lea eax, dword ptr [ebp-24]
'00723e50    50                      push eax
'00723e51    8d4de0                  lea ecx, dword ptr [ebp-20]
'00723e54    51                      push ecx
'00723e55    8d55e4                  lea edx, dword ptr [ebp-1c]
'00723e58    52                      push edx
'00723e59    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00723e5b    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4584, -4588, -4592)
'00723e61    8b45e8                  mov eax, dword ptr [ebp-18]
'00723e64    83c410                  add esp, 10
'00723e67    50                      push eax
'00723e68    685c794500              push 0045795c
'00723e6d    ffd7                    call edi
    var_2092 = (var_2069) & ("Elfes : ")
'00723e6f    8bd0                    mov edx, eax
'00723e71    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00723e74    ffd6                    call esi
'00723e76    db03                    fild dword ptr [ebx]
'00723e78    50                      push eax
'00723e79    dd5d98                  fstp qword ptr [ebp-68]
    'var_130 = (00)
'00723e7c    dd4598                  fld qword ptr [ebp-68]
'00723e7f    dc0d58704000            fmul qword ptr [00407058]
'00723e85    dfe0                    fnstsw ax
'00723e87    a80d                    test al, 0d
'00723e89    0f85f1050000            jne 724480

' *** Reference to "__vbaFPInt"
'00723e8f    ff15fc124000            call dword ptr [004012fc]
'00723e95    83ec08                  sub esp, 08
'00723e98    dd1c24                  fstp qword ptr [esp]
    'var_134 = (00)

' *** Reference to "__vbaStrR8"
'00723e9b    ff1580114000            call dword ptr [00401180]
'00723ea1    8bd0                    mov edx, eax
'00723ea3    8d4de0                  lea ecx, dword ptr [ebp-20]
'00723ea6    ffd6                    call esi
'00723ea8    50                      push eax
'00723ea9    ffd7                    call edi
    var_2461 = (var_2092) & (CStr(Int((((arg_1)) * 0.05#))))
'00723eab    8bd0                    mov edx, eax
'00723ead    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00723eb0    ffd6                    call esi
'00723eb2    50                      push eax
'00723eb3    6870084300              push 00430870
'00723eb8    ffd7                    call edi
    var_870 = (var_2461) & (vbCrLf)
'00723eba    8bd0                    mov edx, eax
'00723ebc    8d4de8                  lea ecx, dword ptr [ebp-18]
'00723ebf    ffd6                    call esi
'00723ec1    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00723ec4    51                      push ecx
'00723ec5    8d55e0                  lea edx, dword ptr [ebp-20]
'00723ec8    52                      push edx
'00723ec9    8d45e4                  lea eax, dword ptr [ebp-1c]
'00723ecc    50                      push eax
'00723ecd    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00723ecf    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4600, -4604, -4608)
'00723ed5    8b4de8                  mov ecx, dword ptr [ebp-18]
'00723ed8    83c410                  add esp, 10
'00723edb    51                      push ecx
'00723edc    68ac794500              push 004579ac
'00723ee1    ffd7                    call edi
    var_40 = (var_870) & ("Nains : ")
'00723ee3    8bd0                    mov edx, eax
'00723ee5    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00723ee8    ffd6                    call esi
'00723eea    db03                    fild dword ptr [ebx]
'00723eec    50                      push eax
'00723eed    dd5d90                  fstp qword ptr [ebp-70]
    'var_8 = (00)
'00723ef0    dd4590                  fld qword ptr [ebp-70]
'00723ef3    dc0dc8704000            fmul qword ptr [004070c8]
'00723ef9    dfe0                    fnstsw ax
'00723efb    a80d                    test al, 0d
'00723efd    0f857d050000            jne 724480

' *** Reference to "__vbaFPInt"
'00723f03    ff15fc124000            call dword ptr [004012fc]
'00723f09    83ec08                  sub esp, 08
'00723f0c    dd1c24                  fstp qword ptr [esp]
    'var_133 = (00)

' *** Reference to "__vbaStrR8"
'00723f0f    ff1580114000            call dword ptr [00401180]
'00723f15    8bd0                    mov edx, eax
'00723f17    8d4de0                  lea ecx, dword ptr [ebp-20]
'00723f1a    ffd6                    call esi
'00723f1c    50                      push eax
'00723f1d    ffd7                    call edi
    var_871 = (var_40) & (CStr(Int((((arg_1)) * 0.03#))))
'00723f1f    8bd0                    mov edx, eax
'00723f21    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00723f24    ffd6                    call esi
'00723f26    50                      push eax
'00723f27    6870084300              push 00430870
'00723f2c    ffd7                    call edi
    var_79 = (var_871) & (vbCrLf)
'00723f2e    8bd0                    mov edx, eax
'00723f30    8d4de8                  lea ecx, dword ptr [ebp-18]
'00723f33    ffd6                    call esi
'00723f35    8d55dc                  lea edx, dword ptr [ebp-24]
'00723f38    52                      push edx
'00723f39    8d45e0                  lea eax, dword ptr [ebp-20]
'00723f3c    50                      push eax
'00723f3d    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00723f40    51                      push ecx
'00723f41    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00723f43    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4616, -4620, -4624)
'00723f49    8b55e8                  mov edx, dword ptr [ebp-18]
'00723f4c    83c410                  add esp, 10
'00723f4f    52                      push edx
'00723f50    68e8774500              push 004577e8
'00723f55    ffd7                    call edi
    var_3 = (var_79) & ("Gnomes : ")
'00723f57    8bd0                    mov edx, eax
'00723f59    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00723f5c    ffd6                    call esi
'00723f5e    db03                    fild dword ptr [ebx]
'00723f60    50                      push eax
'00723f61    dd5d88                  fstp qword ptr [ebp-78]
    'var_131 = (00)
'00723f64    dd4588                  fld qword ptr [ebp-78]
'00723f67    dc0d30714000            fmul qword ptr [00407130]
'00723f6d    dfe0                    fnstsw ax
'00723f6f    a80d                    test al, 0d
'00723f71    0f8509050000            jne 724480

' *** Reference to "__vbaFPInt"
'00723f77    ff15fc124000            call dword ptr [004012fc]
'00723f7d    83ec08                  sub esp, 08
'00723f80    dd1c24                  fstp qword ptr [esp]
    'var_132 = (00)

' *** Reference to "__vbaStrR8"
'00723f83    ff1580114000            call dword ptr [00401180]
'00723f89    8bd0                    mov edx, eax
'00723f8b    8d4de0                  lea ecx, dword ptr [ebp-20]
'00723f8e    ffd6                    call esi
'00723f90    50                      push eax
'00723f91    ffd7                    call edi
    var_67 = (var_3) & (CStr(Int((((arg_1)) * 0.02#))))
'00723f93    8bd0                    mov edx, eax
'00723f95    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00723f98    ffd6                    call esi
'00723f9a    50                      push eax
'00723f9b    6870084300              push 00430870
'00723fa0    ffd7                    call edi
    var_68 = (var_67) & (vbCrLf)
'00723fa2    8bd0                    mov edx, eax
'00723fa4    8d4de8                  lea ecx, dword ptr [ebp-18]
'00723fa7    ffd6                    call esi
'00723fa9    8d45dc                  lea eax, dword ptr [ebp-24]
'00723fac    50                      push eax
'00723fad    8d4de0                  lea ecx, dword ptr [ebp-20]
'00723fb0    51                      push ecx
'00723fb1    8d55e4                  lea edx, dword ptr [ebp-1c]
'00723fb4    52                      push edx
'00723fb5    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00723fb7    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4632, -4636, -4640)
'00723fbd    8b45e8                  mov eax, dword ptr [ebp-18]
'00723fc0    83c410                  add esp, 10
'00723fc3    50                      push eax
'00723fc4    680c544500              push 0045540c
'00723fc9    ffd7                    call edi
    var_2062 = (var_68) & ("Demi-elfes : ")
'00723fcb    8bd0                    mov edx, eax
'00723fcd    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00723fd0    ffd6                    call esi
'00723fd2    db03                    fild dword ptr [ebx]
'00723fd4    50                      push eax
'00723fd5    dd5d80                  fstp qword ptr [ebp-80]
    'var_18 = (00)
'00723fd8    dd4580                  fld qword ptr [ebp-80]
'00723fdb    dc0d28714000            fmul qword ptr [00407128]
'00723fe1    dfe0                    fnstsw ax
'00723fe3    a80d                    test al, 0d
'00723fe5    0f8595040000            jne 724480

' *** Reference to "__vbaFPInt"
'00723feb    ff15fc124000            call dword ptr [004012fc]
'00723ff1    83ec08                  sub esp, 08
'00723ff4    dd1c24                  fstp qword ptr [esp]
    'var_87 = (00)

' *** Reference to "__vbaStrR8"
'00723ff7    ff1580114000            call dword ptr [00401180]
'00723ffd    8bd0                    mov edx, eax
'00723fff    8d4de0                  lea ecx, dword ptr [ebp-20]
'00724002    ffd6                    call esi
'00724004    50                      push eax
'00724005    ffd7                    call edi
    var_70 = (vbNullString) & (CStr(Int((((arg_1)) * 0.01#))))
'00724007    8bd0                    mov edx, eax
'00724009    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0072400c    ffd6                    call esi
'0072400e    50                      push eax
'0072400f    6870084300              push 00430870
'00724014    ffd7                    call edi
    var_872 = (var_70) & (vbCrLf)
'00724016    8bd0                    mov edx, eax
'00724018    8d4de8                  lea ecx, dword ptr [ebp-18]
'0072401b    ffd6                    call esi
'0072401d    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00724020    51                      push ecx
'00724021    8d55e0                  lea edx, dword ptr [ebp-20]
'00724024    52                      push edx
'00724025    8d45e4                  lea eax, dword ptr [ebp-1c]
'00724028    50                      push eax
'00724029    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'0072402b    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4648, -4652, -4656)
'00724031    8b4de8                  mov ecx, dword ptr [ebp-18]
'00724034    83c410                  add esp, 10
'00724037    51                      push ecx
'00724038    6810054500              push 00450510
'0072403d    ffd7                    call edi
    var_40 = (var_872) & ("Demi-orques : ")
'0072403f    8bd0                    mov edx, eax
'00724041    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00724044    ffd6                    call esi
'00724046    db03                    fild dword ptr [ebx]
'00724048    50                      push eax
'00724049    dd9d78ffffff            fstp qword ptr [ebp+ffffff78]
    'var_87 = (00)
'0072404f    dd8578ffffff            fld qword ptr [ebp+ffffff78]
'00724055    dc0d28714000            fmul qword ptr [00407128]
'0072405b    dfe0                    fnstsw ax
'0072405d    a80d                    test al, 0d
'0072405f    0f851b040000            jne 724480

' *** Reference to "__vbaFPInt"
'00724065    ff15fc124000            call dword ptr [004012fc]
'0072406b    83ec08                  sub esp, 08
'0072406e    dd1c24                  fstp qword ptr [esp]
    'var_131 = (00)

' *** Reference to "__vbaStrR8"
'00724071    ff1580114000            call dword ptr [00401180]
'00724077    8bd0                    mov edx, eax
'00724079    8d4de0                  lea ecx, dword ptr [ebp-20]
'0072407c    ffd6                    call esi
'0072407e    50                      push eax
'0072407f    ffd7                    call edi
    var_2583 = (var_40) & (CStr(Int((((arg_1)) * 0.01#))))
'00724081    8bd0                    mov edx, eax
'00724083    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00724086    ffd6                    call esi
'00724088    50                      push eax
'00724089    6870084300              push 00430870
'0072408e    ffd7                    call edi
    var_875 = (var_2583) & (vbCrLf)
'00724090    8bd0                    mov edx, eax
'00724092    8d4de8                  lea ecx, dword ptr [ebp-18]
'00724095    ffd6                    call esi
'00724097    8d55dc                  lea edx, dword ptr [ebp-24]
'0072409a    52                      push edx
'0072409b    8d45e0                  lea eax, dword ptr [ebp-20]
'0072409e    50                      push eax
'0072409f    8d4de4                  lea ecx, dword ptr [ebp-1c]
'007240a2    51                      push ecx
'007240a3    e96d030000              jmp 724415
    
End If
'007240a8    8b55d8                  mov edx, dword ptr [ebp-28]
'007240ab    52                      push edx
'007240ac    68e4554500              push 004555e4
'007240b1    ffd6                    call esi
'007240b3    85c0                    test ax, ax
'007240b5    0f8565030000            jne 724420

If (((arg_0) = ("Intégrée"))) Then
'007240bb    8b45e8                  mov eax, dword ptr [ebp-18]

' *** Reference to "__vbaStrCat"
'007240be    8b3d70104000            mov edi, dword ptr [00401070]
'007240c4    50                      push eax
'007240c5    6820794500              push 00457920
'007240ca    ffd7                    call edi
    var_876 = (var_875) & ("Humains : ")

' *** Reference to "__vbaStrMove"
'007240cc    8b35d0124000            mov esi, dword ptr [004012d0]
'007240d2    8bd0                    mov edx, eax
'007240d4    8d4de4                  lea ecx, dword ptr [ebp-1c]
'007240d7    ffd6                    call esi
'007240d9    8b5d10                  mov ebx, dword ptr [ebp+10]
'007240dc    db03                    fild dword ptr [ebx]
'007240de    50                      push eax
'007240df    dd9d70ffffff            fstp qword ptr [ebp+ffffff70]
    'var_19 = (00)
'007240e5    dd8570ffffff            fld qword ptr [ebp+ffffff70]
'007240eb    dc0d10714000            fmul qword ptr [00407110]
'007240f1    dfe0                    fnstsw ax
'007240f3    a80d                    test al, 0d
'007240f5    0f8585030000            jne 724480

' *** Reference to "__vbaFPInt"
'007240fb    ff15fc124000            call dword ptr [004012fc]
'00724101    83ec08                  sub esp, 08
'00724104    dd1c24                  fstp qword ptr [esp]
    'var_136 = (00)

' *** Reference to "__vbaStrR8"
'00724107    ff1580114000            call dword ptr [00401180]
'0072410d    8bd0                    mov edx, eax
'0072410f    8d4de0                  lea ecx, dword ptr [ebp-20]
'00724112    ffd6                    call esi
'00724114    50                      push eax
'00724115    ffd7                    call edi
    var_2475 = (var_876) & (CStr(Int((((arg_1)) * 0.37#))))
'00724117    8bd0                    mov edx, eax
'00724119    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0072411c    ffd6                    call esi
'0072411e    50                      push eax
'0072411f    6870084300              push 00430870
'00724124    ffd7                    call edi
    var_879 = (var_2475) & (vbCrLf)
'00724126    8bd0                    mov edx, eax
'00724128    8d4de8                  lea ecx, dword ptr [ebp-18]
'0072412b    ffd6                    call esi
'0072412d    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00724130    51                      push ecx
'00724131    8d55e0                  lea edx, dword ptr [ebp-20]
'00724134    52                      push edx
'00724135    8d45e4                  lea eax, dword ptr [ebp-1c]
'00724138    50                      push eax
'00724139    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'0072413b    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4684, -4688, -4692)
'00724141    8b4de8                  mov ecx, dword ptr [ebp-18]
'00724144    83c410                  add esp, 10
'00724147    51                      push ecx
'00724148    683c794500              push 0045793c
'0072414d    ffd7                    call edi
    var_40 = (var_879) & ("Halfelins : ")
'0072414f    8bd0                    mov edx, eax
'00724151    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00724154    ffd6                    call esi
'00724156    db03                    fild dword ptr [ebx]
'00724158    50                      push eax
'00724159    dd9d68ffffff            fstp qword ptr [ebp+ffffff68]
    'var_132 = (00)
'0072415f    dd8568ffffff            fld qword ptr [ebp+ffffff68]
'00724165    dc0d48704000            fmul qword ptr [00407048]
'0072416b    dfe0                    fnstsw ax
'0072416d    a80d                    test al, 0d
'0072416f    0f850b030000            jne 724480

' *** Reference to "__vbaFPInt"
'00724175    ff15fc124000            call dword ptr [004012fc]
'0072417b    83ec08                  sub esp, 08
'0072417e    dd1c24                  fstp qword ptr [esp]
    'var_135 = (00)

' *** Reference to "__vbaStrR8"
'00724181    ff1580114000            call dword ptr [00401180]
'00724187    8bd0                    mov edx, eax
'00724189    8d4de0                  lea ecx, dword ptr [ebp-20]
'0072418c    ffd6                    call esi
'0072418e    50                      push eax
'0072418f    ffd7                    call edi
    var_881 = (var_40) & (CStr(Int((((arg_1)) * 0.2#))))
'00724191    8bd0                    mov edx, eax
'00724193    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00724196    ffd6                    call esi
'00724198    50                      push eax
'00724199    6870084300              push 00430870
'0072419e    ffd7                    call edi
    var_882 = (var_881) & (vbCrLf)
'007241a0    8bd0                    mov edx, eax
'007241a2    8d4de8                  lea ecx, dword ptr [ebp-18]
'007241a5    ffd6                    call esi
'007241a7    8d55dc                  lea edx, dword ptr [ebp-24]
'007241aa    52                      push edx
'007241ab    8d45e0                  lea eax, dword ptr [ebp-20]
'007241ae    50                      push eax
'007241af    8d4de4                  lea ecx, dword ptr [ebp-1c]
'007241b2    51                      push ecx
'007241b3    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'007241b5    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4700, -4704, -4708)
'007241bb    8b55e8                  mov edx, dword ptr [ebp-18]
'007241be    83c410                  add esp, 10
'007241c1    52                      push edx
'007241c2    685c794500              push 0045795c
'007241c7    ffd7                    call edi
    var_3 = (var_882) & ("Elfes : ")
'007241c9    8bd0                    mov edx, eax
'007241cb    8d4de4                  lea ecx, dword ptr [ebp-1c]
'007241ce    ffd6                    call esi
'007241d0    db03                    fild dword ptr [ebx]
'007241d2    50                      push eax
'007241d3    dd9d60ffffff            fstp qword ptr [ebp+ffffff60]
    'var_20 = (00)
'007241d9    dd8560ffffff            fld qword ptr [ebp+ffffff60]
'007241df    dc0d08714000            fmul qword ptr [00407108]
'007241e5    dfe0                    fnstsw ax
'007241e7    a80d                    test al, 0d
'007241e9    0f8591020000            jne 724480

' *** Reference to "__vbaFPInt"
'007241ef    ff15fc124000            call dword ptr [004012fc]
'007241f5    83ec08                  sub esp, 08
'007241f8    dd1c24                  fstp qword ptr [esp]
    'var_134 = (00)

' *** Reference to "__vbaStrR8"
'007241fb    ff1580114000            call dword ptr [00401180]
'00724201    8bd0                    mov edx, eax
'00724203    8d4de0                  lea ecx, dword ptr [ebp-20]
'00724206    ffd6                    call esi
'00724208    50                      push eax
'00724209    ffd7                    call edi
    var_884 = (var_3) & (CStr(Int((((arg_1)) * 0.18#))))
'0072420b    8bd0                    mov edx, eax
'0072420d    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00724210    ffd6                    call esi
'00724212    50                      push eax
'00724213    6870084300              push 00430870
'00724218    ffd7                    call edi
    var_2445 = (var_884) & (vbCrLf)
'0072421a    8bd0                    mov edx, eax
'0072421c    8d4de8                  lea ecx, dword ptr [ebp-18]
'0072421f    ffd6                    call esi
'00724221    8d45dc                  lea eax, dword ptr [ebp-24]
'00724224    50                      push eax
'00724225    8d4de0                  lea ecx, dword ptr [ebp-20]
'00724228    51                      push ecx
'00724229    8d55e4                  lea edx, dword ptr [ebp-1c]
'0072422c    52                      push edx
'0072422d    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'0072422f    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4716, -4720, -4724)
'00724235    8b45e8                  mov eax, dword ptr [ebp-18]
'00724238    83c410                  add esp, 10
'0072423b    50                      push eax
'0072423c    68ac794500              push 004579ac
'00724241    ffd7                    call edi
    var_886 = (var_2445) & ("Nains : ")
'00724243    8bd0                    mov edx, eax
'00724245    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00724248    ffd6                    call esi
'0072424a    db03                    fild dword ptr [ebx]
'0072424c    50                      push eax
'0072424d    dd9d58ffffff            fstp qword ptr [ebp+ffffff58]
    'var_133 = (00)
'00724253    dd8558ffffff            fld qword ptr [ebp+ffffff58]
'00724259    dc0d50704000            fmul qword ptr [00407050]
'0072425f    dfe0                    fnstsw ax
'00724261    a80d                    test al, 0d
'00724263    0f8517020000            jne 724480

' *** Reference to "__vbaFPInt"
'00724269    ff15fc124000            call dword ptr [004012fc]
'0072426f    83ec08                  sub esp, 08
'00724272    dd1c24                  fstp qword ptr [esp]
    'var_133 = (00)

' *** Reference to "__vbaStrR8"
'00724275    ff1580114000            call dword ptr [00401180]
'0072427b    8bd0                    mov edx, eax
'0072427d    8d4de0                  lea ecx, dword ptr [ebp-20]
'00724280    ffd6                    call esi
'00724282    50                      push eax
'00724283    ffd7                    call edi
    var_1470 = (var_886) & (CStr(Int((((arg_1)) * 0.1#))))
'00724285    8bd0                    mov edx, eax
'00724287    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0072428a    ffd6                    call esi
'0072428c    50                      push eax
'0072428d    6870084300              push 00430870
'00724292    ffd7                    call edi
    var_2369 = (var_1470) & (vbCrLf)
'00724294    8bd0                    mov edx, eax
'00724296    8d4de8                  lea ecx, dword ptr [ebp-18]
'00724299    ffd6                    call esi
'0072429b    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0072429e    51                      push ecx
'0072429f    8d55e0                  lea edx, dword ptr [ebp-20]
'007242a2    52                      push edx
'007242a3    8d45e4                  lea eax, dword ptr [ebp-1c]
'007242a6    50                      push eax
'007242a7    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'007242a9    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4732, -4736, -4740)
'007242af    8b4de8                  mov ecx, dword ptr [ebp-18]
'007242b2    83c410                  add esp, 10
'007242b5    51                      push ecx
'007242b6    68e8774500              push 004577e8
'007242bb    ffd7                    call edi
    var_40 = (var_2369) & ("Gnomes : ")
'007242bd    8bd0                    mov edx, eax
'007242bf    8d4de4                  lea ecx, dword ptr [ebp-1c]
'007242c2    ffd6                    call esi
'007242c4    db03                    fild dword ptr [ebx]
'007242c6    50                      push eax
'007242c7    dd9d50ffffff            fstp qword ptr [ebp+ffffff50]
    'var_21 = (00)
'007242cd    dd8550ffffff            fld qword ptr [ebp+ffffff50]
'007242d3    dc0d00714000            fmul qword ptr [00407100]
'007242d9    dfe0                    fnstsw ax
'007242db    a80d                    test al, 0d
'007242dd    0f859d010000            jne 724480

' *** Reference to "__vbaFPInt"
'007242e3    ff15fc124000            call dword ptr [004012fc]
'007242e9    83ec08                  sub esp, 08
'007242ec    dd1c24                  fstp qword ptr [esp]
    'var_132 = (00)

' *** Reference to "__vbaStrR8"
'007242ef    ff1580114000            call dword ptr [00401180]
'007242f5    8bd0                    mov edx, eax
'007242f7    8d4de0                  lea ecx, dword ptr [ebp-20]
'007242fa    ffd6                    call esi
'007242fc    50                      push eax
'007242fd    ffd7                    call edi
    var_1503 = (var_40) & (CStr(Int((((arg_1)) * 0.07#))))
'007242ff    8bd0                    mov edx, eax
'00724301    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00724304    ffd6                    call esi
'00724306    50                      push eax
'00724307    6870084300              push 00430870
'0072430c    ffd7                    call edi
    var_2444 = (var_1503) & (vbCrLf)
'0072430e    8bd0                    mov edx, eax
'00724310    8d4de8                  lea ecx, dword ptr [ebp-18]
'00724313    ffd6                    call esi
'00724315    8d55dc                  lea edx, dword ptr [ebp-24]
'00724318    52                      push edx
'00724319    8d45e0                  lea eax, dword ptr [ebp-20]
'0072431c    50                      push eax
'0072431d    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00724320    51                      push ecx
'00724321    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00724323    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4748, -4752, -4756)
'00724329    8b55e8                  mov edx, dword ptr [ebp-18]
'0072432c    83c410                  add esp, 10
'0072432f    52                      push edx
'00724330    680c544500              push 0045540c
'00724335    ffd7                    call edi
    var_3 = (var_2444) & ("Demi-elfes : ")
'00724337    8bd0                    mov edx, eax
'00724339    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0072433c    ffd6                    call esi
'0072433e    db03                    fild dword ptr [ebx]
'00724340    50                      push eax
'00724341    dd9d48ffffff            fstp qword ptr [ebp+ffffff48]
    'var_134 = (00)
'00724347    dd8548ffffff            fld qword ptr [ebp+ffffff48]
'0072434d    dc0d58704000            fmul qword ptr [00407058]
'00724353    dfe0                    fnstsw ax
'00724355    a80d                    test al, 0d
'00724357    0f8523010000            jne 724480

' *** Reference to "__vbaFPInt"
'0072435d    ff15fc124000            call dword ptr [004012fc]
'00724363    83ec08                  sub esp, 08
'00724366    dd1c24                  fstp qword ptr [esp]
    'var_87 = (00)

' *** Reference to "__vbaStrR8"
'00724369    ff1580114000            call dword ptr [00401180]
'0072436f    8bd0                    mov edx, eax
'00724371    8d4de0                  lea ecx, dword ptr [ebp-20]
'00724374    ffd6                    call esi
'00724376    50                      push eax
'00724377    ffd7                    call edi
    var_2189 = (var_3) & (CStr(Int((((arg_1)) * 0.05#))))
'00724379    8bd0                    mov edx, eax
'0072437b    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0072437e    ffd6                    call esi
'00724380    50                      push eax
'00724381    6870084300              push 00430870
'00724386    ffd7                    call edi
    var_1528 = (var_2189) & (vbCrLf)
'00724388    8bd0                    mov edx, eax
'0072438a    8d4de8                  lea ecx, dword ptr [ebp-18]
'0072438d    ffd6                    call esi
'0072438f    8d45dc                  lea eax, dword ptr [ebp-24]
'00724392    50                      push eax
'00724393    8d4de0                  lea ecx, dword ptr [ebp-20]
'00724396    51                      push ecx
'00724397    8d55e4                  lea edx, dword ptr [ebp-1c]
'0072439a    52                      push edx
'0072439b    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'0072439d    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4764, -4768, -4772)
'007243a3    8b45e8                  mov eax, dword ptr [ebp-18]
'007243a6    83c410                  add esp, 10
'007243a9    50                      push eax
'007243aa    6810054500              push 00450510
'007243af    ffd7                    call edi
    var_2124 = (var_1528) & ("Demi-orques : ")
'007243b1    8bd0                    mov edx, eax
'007243b3    8d4de4                  lea ecx, dword ptr [ebp-1c]
'007243b6    ffd6                    call esi
'007243b8    db03                    fild dword ptr [ebx]
'007243ba    50                      push eax
'007243bb    dd9d40ffffff            fstp qword ptr [ebp+ffffff40]
    'var_22 = (00)
'007243c1    dd8540ffffff            fld qword ptr [ebp+ffffff40]
'007243c7    dc0dc8704000            fmul qword ptr [004070c8]
'007243cd    dfe0                    fnstsw ax
'007243cf    a80d                    test al, 0d
'007243d1    0f85a9000000            jne 724480

' *** Reference to "__vbaFPInt"
'007243d7    ff15fc124000            call dword ptr [004012fc]
'007243dd    83ec08                  sub esp, 08
'007243e0    dd1c24                  fstp qword ptr [esp]
    'var_131 = (00)

' *** Reference to "__vbaStrR8"
'007243e3    ff1580114000            call dword ptr [00401180]
'007243e9    8bd0                    mov edx, eax
'007243eb    8d4de0                  lea ecx, dword ptr [ebp-20]
'007243ee    ffd6                    call esi
'007243f0    50                      push eax
'007243f1    ffd7                    call edi
    var_2446 = (var_2124) & (CStr(Int((((arg_1)) * 0.03#))))
'007243f3    8bd0                    mov edx, eax
'007243f5    8d4ddc                  lea ecx, dword ptr [ebp-24]
'007243f8    ffd6                    call esi
'007243fa    50                      push eax
'007243fb    6870084300              push 00430870
'00724400    ffd7                    call edi
    var_1532 = (var_2446) & (vbCrLf)
'00724402    8bd0                    mov edx, eax
'00724404    8d4de8                  lea ecx, dword ptr [ebp-18]
'00724407    ffd6                    call esi
'00724409    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0072440c    51                      push ecx
'0072440d    8d55e0                  lea edx, dword ptr [ebp-20]
'00724410    52                      push edx
'00724411    8d45e4                  lea eax, dword ptr [ebp-1c]
'00724414    50                      push eax
'00724415    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00724417    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( -4780, -4784, -4788)
'0072441d    83c410                  add esp, 10
    
End If
'00724420    9b                      fwait
'00724421    6859447200              push 00724459
'00724426    eb27                    jmp 72444f
'00724428    f645fc04                test byte ptr [ebp-04], 04
'0072442c    7409                    je 724437
'0072442e    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeStr"
'00724431    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_41)
'00724437    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0072443a    51                      push ecx
'0072443b    8d55e0                  lea edx, dword ptr [ebp-20]
'0072443e    52                      push edx
'0072443f    8d45e4                  lea eax, dword ptr [ebp-1c]
'00724442    50                      push eax
'00724443    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'00724445    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4780, -4784, -4788)
'0072444b    83c410                  add esp, 10
'0072444e    c3                      ret
'0072444f    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeStr"
'00724452    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_45)
'00724458    c3                      ret
'00724459    8b4508                  mov eax, dword ptr [ebp+08]
'0072445c    8b08                    mov ecx, dword ptr [eax]
'0072445e    50                      push eax
'0072445f    ff5108                  call dword ptr [ecx+08]
'00724462    8b45e8                  mov eax, dword ptr [ebp-18]
'00724465    8b5514                  mov edx, dword ptr [ebp+14]
'00724468    8902                    mov dword ptr [edx], eax
'0072446a    8b45fc                  mov eax, dword ptr [ebp-04]
'0072446d    8b4dec                  mov ecx, dword ptr [ebp-14]
'00724470    5f                      pop edi
'00724471    5e                      pop esi
    'Reference to '__except_list'
'00724472    64890d00000000          mov dword ptr fs:[00000000], ecx
'00724479    5b                      pop ebx
'0072447a    8be5                    mov esp, ebp
'0072447c    5d                      pop ebp
'0072447d    c21000                  ret 0010


End Function


'Event for cmbPrint
Private Sub cmbPrint_Click()
'0071f150    55                      push ebp
'0071f151    8bec                    mov ebp, esp
'0071f153    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0071f156    6866724000              push 00407266
'0071f15b    64a100000000            mov ax, word ptr fs:[00000000]
'0071f161    50                      push eax
    'Reference to '__except_list'
'0071f162    64892500000000          mov dword ptr fs:[00000000], esp
'0071f169    83ec40                  sub esp, 40
'0071f16c    53                      push ebx
'0071f16d    56                      push esi
'0071f16e    57                      push edi
'0071f16f    8965f4                  mov dword ptr [ebp-0c], esp
'0071f172    c745f8a86f4000          mov dword ptr [ebp-08], 00406fa8
'0071f179    8b4508                  mov eax, dword ptr [ebp+08]
'0071f17c    8bc8                    mov ecx, eax
'0071f17e    83e101                  and ecx, 01
'0071f181    894dfc                  mov dword ptr [ebp-04], ecx
'0071f184    83e0fe                  and eax, fffffffe
'0071f187    8b10                    mov edx, dword ptr [eax]
'0071f189    50                      push eax
'0071f18a    894508                  mov dword ptr [ebp+08], eax
'0071f18d    ff5204                  call dword ptr [edx+04]
'0071f190    a124be7200              mov ax, word ptr [0072be24]
'0071f195    33ff                    xor edi, edi
'0071f197    83cbff                  or ebx, ffffffff
'0071f19a    3bc7                    cmp eax, edi
'0071f19c    897de8                  mov dword ptr [ebp-18], edi
'0071f19f    897de4                  mov dword ptr [ebp-1c], edi
'0071f1a2    897de0                  mov dword ptr [ebp-20], edi
'0071f1a5    7510                    jne 71f1b7

If (0 = 0) Then
'0071f1a7    6824be7200              push 0072be24
'0071f1ac    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0071f1b1    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
End If
'0071f1b7    8b3524be7200            mov esi, dword ptr [0072be24]
'0071f1bd    8b06                    mov eax, dword ptr [esi]
'0071f1bf    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071f1c2    51                      push ecx
'0071f1c3    56                      push esi
'0071f1c4    ff5020                  call dword ptr [eax+20]
Set var_41 = var_2.Printer
'0071f1c7    dbe2                    fnclex
'0071f1c9    3bc7                    cmp eax, edi
'0071f1cb    7d0f                    jge 71f1dc

If (var_2.Printer < 0) Then
'0071f1cd    6a20                    push 20
'0071f1cf    6860004300              push 00430060
'0071f1d4    56                      push esi
'0071f1d5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071f1d6    ff1580104000            call dword ptr [00401080]
    
End If
'0071f1dc    8b75d4                  mov esi, dword ptr [ebp-2c]
'0071f1df    8b7ddc                  mov edi, dword ptr [ebp-24]
'0071f1e2    83ec10                  sub esp, 10
'0071f1e5    8bd4                    mov edx, esp
'0071f1e7    b80b000000              mov eax, 0000000b
'0071f1ec    8902                    mov dword ptr [edx], eax
'0071f1ee    8b45e8                  mov eax, dword ptr [ebp-18]
'0071f1f1    897204                  mov dword ptr [edx+04], esi
'0071f1f4    895a08                  mov dword ptr [edx+08], ebx

' *** Reference to "__vbaLateIdSt"
'0071f1f7    8b1dec124000            mov ebx, dword ptr [004012ec]
'0071f1fd    6826000100              push 00010026
'0071f202    50                      push eax
'0071f203    897a0c                  mov dword ptr [edx+0c], edi
'0071f206    ffd3                    call ebx
var_41.[UNMANAGED] = -1
'0071f208    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0071f20b    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0071f211    a124be7200              mov ax, word ptr [0072be24]
'0071f216    85c0                    test ax, ax
'0071f218    7510                    jne 71f22a
'0071f21a    6824be7200              push 0072be24
'0071f21f    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0071f224    ff152c124000            call dword ptr [0040122c]
Set var_2 = New Global
'0071f22a    a124be7200              mov ax, word ptr [0072be24]
'0071f22f    8b08                    mov ecx, dword ptr [eax]
'0071f231    8d55e8                  lea edx, dword ptr [ebp-18]
'0071f234    52                      push edx
'0071f235    50                      push eax
'0071f236    8945bc                  mov dword ptr [ebp-44], eax
'0071f239    ff5120                  call dword ptr [ecx+20]
Set var_41 = var_2.Printer
'0071f23c    dbe2                    fnclex
'0071f23e    85c0                    test ax, ax
'0071f240    7d12                    jge 71f254
'0071f242    8b4dbc                  mov ecx, dword ptr [ebp-44]
'0071f245    6a20                    push 20
'0071f247    6860004300              push 00430060
'0071f24c    51                      push ecx
'0071f24d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071f24e    ff1580104000            call dword ptr [00401080]
'0071f254    83ec10                  sub esp, 10
'0071f257    8bd4                    mov edx, esp
'0071f259    b802000000              mov eax, 00000002
'0071f25e    8902                    mov dword ptr [edx], eax
'0071f260    897204                  mov dword ptr [edx+04], esi
'0071f263    b806000000              mov eax, 00000006
'0071f268    894208                  mov dword ptr [edx+08], eax
'0071f26b    8b45e8                  mov eax, dword ptr [ebp-18]
'0071f26e    6810000100              push 00010010
'0071f273    50                      push eax
'0071f274    897a0c                  mov dword ptr [edx+0c], edi
'0071f277    ffd3                    call ebx
var_41.[UNMANAGED] = 6
'0071f279    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0071f27c    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0071f282    a124be7200              mov ax, word ptr [0072be24]
'0071f287    85c0                    test ax, ax
'0071f289    7510                    jne 71f29b
'0071f28b    6824be7200              push 0072be24
'0071f290    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0071f295    ff152c124000            call dword ptr [0040122c]
Set var_2 = New Global
'0071f29b    a124be7200              mov ax, word ptr [0072be24]
'0071f2a0    8b08                    mov ecx, dword ptr [eax]
'0071f2a2    8d55e8                  lea edx, dword ptr [ebp-18]
'0071f2a5    52                      push edx
'0071f2a6    50                      push eax
'0071f2a7    8945bc                  mov dword ptr [ebp-44], eax
'0071f2aa    ff5120                  call dword ptr [ecx+20]
Set var_41 = var_2.Printer
'0071f2ad    dbe2                    fnclex
'0071f2af    85c0                    test ax, ax
'0071f2b1    7d12                    jge 71f2c5
'0071f2b3    8b4dbc                  mov ecx, dword ptr [ebp-44]
'0071f2b6    6a20                    push 20
'0071f2b8    6860004300              push 00430060
'0071f2bd    51                      push ecx
'0071f2be    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071f2bf    ff1580104000            call dword ptr [00401080]
'0071f2c5    83ec10                  sub esp, 10
'0071f2c8    8bd4                    mov edx, esp
'0071f2ca    b802000000              mov eax, 00000002
'0071f2cf    8902                    mov dword ptr [edx], eax
'0071f2d1    897204                  mov dword ptr [edx+04], esi
'0071f2d4    b801000000              mov eax, 00000001
'0071f2d9    894208                  mov dword ptr [edx+08], eax
'0071f2dc    8b45e8                  mov eax, dword ptr [ebp-18]
'0071f2df    681f000100              push 0001001f
'0071f2e4    50                      push eax
'0071f2e5    897a0c                  mov dword ptr [edx+0c], edi
'0071f2e8    ffd3                    call ebx
var_41.[UNMANAGED] = 1
'0071f2ea    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0071f2ed    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0071f2f3    a124be7200              mov ax, word ptr [0072be24]
'0071f2f8    85c0                    test ax, ax
'0071f2fa    7510                    jne 71f30c
'0071f2fc    6824be7200              push 0072be24
'0071f301    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0071f306    ff152c124000            call dword ptr [0040122c]
Set var_2 = New Global
'0071f30c    a124be7200              mov ax, word ptr [0072be24]
'0071f311    8b08                    mov ecx, dword ptr [eax]
'0071f313    8d55e8                  lea edx, dword ptr [ebp-18]
'0071f316    52                      push edx
'0071f317    50                      push eax
'0071f318    8945bc                  mov dword ptr [ebp-44], eax
'0071f31b    ff5120                  call dword ptr [ecx+20]
Set var_41 = var_2.Printer
'0071f31e    dbe2                    fnclex
'0071f320    85c0                    test ax, ax
'0071f322    7d12                    jge 71f336
'0071f324    8b4dbc                  mov ecx, dword ptr [ebp-44]
'0071f327    6a20                    push 20
'0071f329    6860004300              push 00430060
'0071f32e    51                      push ecx
'0071f32f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071f330    ff1580104000            call dword ptr [00401080]
'0071f336    83ec10                  sub esp, 10
'0071f339    8bd4                    mov edx, esp
'0071f33b    b802000000              mov eax, 00000002
'0071f340    8902                    mov dword ptr [edx], eax
'0071f342    897204                  mov dword ptr [edx+04], esi
'0071f345    b809000000              mov eax, 00000009
'0071f34a    894208                  mov dword ptr [edx+08], eax
'0071f34d    8b45e8                  mov eax, dword ptr [ebp-18]
'0071f350    6820000100              push 00010020
'0071f355    50                      push eax
'0071f356    897a0c                  mov dword ptr [edx+0c], edi
'0071f359    ffd3                    call ebx
var_41.[UNMANAGED] = 9
'0071f35b    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0071f35e    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0071f364    a124be7200              mov ax, word ptr [0072be24]
'0071f369    85c0                    test ax, ax
'0071f36b    7510                    jne 71f37d
'0071f36d    6824be7200              push 0072be24
'0071f372    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0071f377    ff152c124000            call dword ptr [0040122c]
Set var_2 = New Global
'0071f37d    a124be7200              mov ax, word ptr [0072be24]
'0071f382    8b08                    mov ecx, dword ptr [eax]
'0071f384    8d55e8                  lea edx, dword ptr [ebp-18]
'0071f387    52                      push edx
'0071f388    50                      push eax
'0071f389    8945bc                  mov dword ptr [ebp-44], eax
'0071f38c    ff5120                  call dword ptr [ecx+20]
Set var_41 = var_2.Printer
'0071f38f    dbe2                    fnclex
'0071f391    85c0                    test ax, ax
'0071f393    7d12                    jge 71f3a7
'0071f395    8b4dbc                  mov ecx, dword ptr [ebp-44]
'0071f398    6a20                    push 20
'0071f39a    6860004300              push 00430060
'0071f39f    51                      push ecx
'0071f3a0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071f3a1    ff1580104000            call dword ptr [00401080]
'0071f3a7    83ec10                  sub esp, 10
'0071f3aa    8bd4                    mov edx, esp
'0071f3ac    b802000000              mov eax, 00000002
'0071f3b1    8902                    mov dword ptr [edx], eax
'0071f3b3    897204                  mov dword ptr [edx+04], esi
'0071f3b6    b8fcffffff              mov eax, fffffffc
'0071f3bb    894208                  mov dword ptr [edx+08], eax
'0071f3be    8b45e8                  mov eax, dword ptr [ebp-18]
'0071f3c1    6827000100              push 00010027
'0071f3c6    50                      push eax
'0071f3c7    897a0c                  mov dword ptr [edx+0c], edi
'0071f3ca    ffd3                    call ebx
var_41.[UNMANAGED] = -4
'0071f3cc    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0071f3cf    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0071f3d5    a124be7200              mov ax, word ptr [0072be24]
'0071f3da    85c0                    test ax, ax
'0071f3dc    c745d801000000          mov dword ptr [ebp-28], 00000001
'0071f3e3    c745d002000000          mov dword ptr [ebp-30], 00000002
'0071f3ea    7510                    jne 71f3fc
'0071f3ec    6824be7200              push 0072be24
'0071f3f1    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0071f3f6    ff152c124000            call dword ptr [0040122c]
Set var_2 = New Global
'0071f3fc    a124be7200              mov ax, word ptr [0072be24]
'0071f401    8b08                    mov ecx, dword ptr [eax]
'0071f403    8d55e8                  lea edx, dword ptr [ebp-18]
'0071f406    52                      push edx
'0071f407    50                      push eax
'0071f408    8945bc                  mov dword ptr [ebp-44], eax
'0071f40b    ff5120                  call dword ptr [ecx+20]
Set var_41 = var_2.Printer
'0071f40e    dbe2                    fnclex
'0071f410    85c0                    test ax, ax
'0071f412    7d12                    jge 71f426
'0071f414    8b4dbc                  mov ecx, dword ptr [ebp-44]
'0071f417    6a20                    push 20
'0071f419    6860004300              push 00430060
'0071f41e    51                      push ecx
'0071f41f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071f420    ff1580104000            call dword ptr [00401080]
'0071f426    8b45d0                  mov eax, dword ptr [ebp-30]
'0071f429    8b4dd8                  mov ecx, dword ptr [ebp-28]
'0071f42c    83ec10                  sub esp, 10
'0071f42f    8bd4                    mov edx, esp
'0071f431    8902                    mov dword ptr [edx], eax
'0071f433    897204                  mov dword ptr [edx+04], esi
'0071f436    894a08                  mov dword ptr [edx+08], ecx
'0071f439    897a0c                  mov dword ptr [edx+0c], edi
'0071f43c    8b55e8                  mov edx, dword ptr [ebp-18]
'0071f43f    6824000100              push 00010024
'0071f444    52                      push edx
'0071f445    ffd3                    call ebx
var_41.[UNMANAGED] = var_45

' *** Reference to "__vbaFreeObj"
'0071f447    8b1d08134000            mov ebx, dword ptr [00401308]
'0071f44d    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071f450    ffd3                    call ebx
'#Cleanup(var_41)
'0071f452    a124be7200              mov ax, word ptr [0072be24]
'0071f457    33ff                    xor edi, edi
var_num7 = Empty
'0071f459    3bc7                    cmp eax, edi
'0071f45b    7510                    jne 71f46d

If (var_2 = 0) Then
'0071f45d    6824be7200              push 0072be24
'0071f462    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0071f467    ff152c124000            call dword ptr [0040122c]
    Set var_2 = New Global
End If
'0071f46d    8b3524be7200            mov esi, dword ptr [0072be24]
'0071f473    8b06                    mov eax, dword ptr [esi]
'0071f475    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071f478    51                      push ecx
'0071f479    56                      push esi
'0071f47a    ff5020                  call dword ptr [eax+20]
Set var_41 = var_2.Printer
'0071f47d    dbe2                    fnclex
'0071f47f    3bc7                    cmp eax, edi
'0071f481    7d0f                    jge 71f492

If (var_2.Printer < 0) Then
'0071f483    6a20                    push 20
'0071f485    6860004300              push 00430060
'0071f48a    56                      push esi
'0071f48b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071f48c    ff1580104000            call dword ptr [00401080]
    
End If
'0071f492    8b4508                  mov eax, dword ptr [ebp+08]
'0071f495    8b10                    mov edx, dword ptr [eax]
'0071f497    50                      push eax
'0071f498    ff9204030000            call dword ptr [edx+00000304]

' *** Reference to "__vbaObjSet"
'0071f49e    8b35b4104000            mov esi, dword ptr [004010b4]
'0071f4a4    50                      push eax
'0071f4a5    8d45e0                  lea eax, dword ptr [ebp-20]
'0071f4a8    50                      push eax
'0071f4a9    ffd6                    call esi
Set var_3 = Me
'0071f4ab    8b45e0                  mov eax, dword ptr [ebp-20]
'0071f4ae    50                      push eax
'0071f4af    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0071f4b2    51                      push ecx
'0071f4b3    897de0                  mov dword ptr [ebp-20], edi
'0071f4b6    ffd6                    call esi
Set var_40 = var_3
'0071f4b8    8b55e8                  mov edx, dword ptr [ebp-18]
'0071f4bb    50                      push eax
'0071f4bc    52                      push edx
'0071f4bd    6854fd4300              push 0043fd54

' *** Reference to "__vbaPrintObj"
'0071f4c2    ff1548114000            call dword ptr [00401148]
var_41.Print var_40 (Object)
'0071f4c8    8d45e0                  lea eax, dword ptr [ebp-20]
'0071f4cb    50                      push eax
'0071f4cc    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0071f4cf    51                      push ecx
'0071f4d0    8d55e8                  lea edx, dword ptr [ebp-18]
'0071f4d3    52                      push edx
'0071f4d4    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0071f4d6    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40, var_3)
'0071f4dc    a124be7200              mov ax, word ptr [0072be24]
'0071f4e1    83c41c                  add esp, 1c
'0071f4e4    3bc7                    cmp eax, edi
'0071f4e6    7510                    jne 71f4f8

If (var_2 = 0) Then
'0071f4e8    6824be7200              push 0072be24
'0071f4ed    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0071f4f2    ff152c124000            call dword ptr [0040122c]
    Set var_2 = New Global
End If
'0071f4f8    8b3524be7200            mov esi, dword ptr [0072be24]
'0071f4fe    8b06                    mov eax, dword ptr [esi]
'0071f500    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071f503    51                      push ecx
'0071f504    56                      push esi
'0071f505    ff5020                  call dword ptr [eax+20]
Set var_41 = var_2.Printer
'0071f508    dbe2                    fnclex
'0071f50a    3bc7                    cmp eax, edi
'0071f50c    7d0f                    jge 71f51d

If (var_2.Printer < 0) Then
'0071f50e    6a20                    push 20
'0071f510    6860004300              push 00430060
'0071f515    56                      push esi
'0071f516    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071f517    ff1580104000            call dword ptr [00401080]
    
End If
'0071f51d    8b55e8                  mov edx, dword ptr [ebp-18]
'0071f520    57                      push edi
'0071f521    6800000200              push 00020000
'0071f526    52                      push edx

' *** Reference to "__vbaLateIdCall"
'0071f527    ff1538104000            call dword ptr [00401038]
Call var_41.EndDoc()
'0071f52d    83c40c                  add esp, 0c
'0071f530    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071f533    ffd3                    call ebx
'#Cleanup(var_41)
'0071f535    897dfc                  mov dword ptr [ebp-04], edi
'0071f538    6858f57100              push 0071f558
'0071f53d    eb18                    jmp 71f557
'0071f53f    8d45e0                  lea eax, dword ptr [ebp-20]
'0071f542    50                      push eax
'0071f543    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0071f546    51                      push ecx
'0071f547    8d55e8                  lea edx, dword ptr [ebp-18]
'0071f54a    52                      push edx
'0071f54b    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0071f54d    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_41, var_40, var_3)
'0071f553    83c410                  add esp, 10
'0071f556    c3                      ret
'0071f557    c3                      ret
'0071f558    8b4508                  mov eax, dword ptr [ebp+08]
'0071f55b    8b08                    mov ecx, dword ptr [eax]
'0071f55d    50                      push eax
'0071f55e    ff5108                  call dword ptr [ecx+08]
'0071f561    8b45fc                  mov eax, dword ptr [ebp-04]
'0071f564    8b4dec                  mov ecx, dword ptr [ebp-14]
'0071f567    5f                      pop edi
'0071f568    5e                      pop esi
    'Reference to '__except_list'
'0071f569    64890d00000000          mov dword ptr fs:[00000000], ecx
'0071f570    5b                      pop ebx
'0071f571    8be5                    mov esp, ebp
'0071f573    5d                      pop ebp
'0071f574    c20400                  ret 0004


End Sub


'Event for cmdGenerer
Private Sub cmdGenerer_Click()
'0071f640    55                      push ebp
'0071f641    8bec                    mov ebp, esp
'0071f643    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0071f646    6866724000              push 00407266
'0071f64b    64a100000000            mov ax, word ptr fs:[00000000]
'0071f651    50                      push eax
    'Reference to '__except_list'
'0071f652    64892500000000          mov dword ptr fs:[00000000], esp
'0071f659    81ec70010000            sub esp, 00000170
'0071f65f    53                      push ebx
'0071f660    56                      push esi
'0071f661    57                      push edi
'0071f662    8965f4                  mov dword ptr [ebp-0c], esp
'0071f665    c745f8c86f4000          mov dword ptr [ebp-08], 00406fc8
'0071f66c    8b7508                  mov esi, dword ptr [ebp+08]
'0071f66f    8bc6                    mov eax, esi
'0071f671    83e001                  and eax, 01
'0071f674    8945fc                  mov dword ptr [ebp-04], eax
'0071f677    83e6fe                  and esi, fffffffe
'0071f67a    8b0e                    mov ecx, dword ptr [esi]
'0071f67c    56                      push esi
'0071f67d    897508                  mov dword ptr [ebp+08], esi
'0071f680    ff5104                  call dword ptr [ecx+04]
'0071f683    8b16                    mov edx, dword ptr [esi]
'0071f685    33db                    xor ebx, ebx
'0071f687    56                      push esi
'0071f688    895de8                  mov dword ptr [ebp-18], ebx
'0071f68b    895de4                  mov dword ptr [ebp-1c], ebx
'0071f68e    895de0                  mov dword ptr [ebp-20], ebx
'0071f691    895ddc                  mov dword ptr [ebp-24], ebx
'0071f694    895dd8                  mov dword ptr [ebp-28], ebx
'0071f697    895dd4                  mov dword ptr [ebp-2c], ebx
'0071f69a    895dd0                  mov dword ptr [ebp-30], ebx
'0071f69d    895dcc                  mov dword ptr [ebp-34], ebx
'0071f6a0    895dc8                  mov dword ptr [ebp-38], ebx
'0071f6a3    895dc4                  mov dword ptr [ebp-3c], ebx
'0071f6a6    895dc0                  mov dword ptr [ebp-40], ebx
'0071f6a9    895dbc                  mov dword ptr [ebp-44], ebx
'0071f6ac    895db8                  mov dword ptr [ebp-48], ebx
'0071f6af    895db4                  mov dword ptr [ebp-4c], ebx
'0071f6b2    895da4                  mov dword ptr [ebp-5c], ebx
'0071f6b5    895d94                  mov dword ptr [ebp-6c], ebx
'0071f6b8    895d84                  mov dword ptr [ebp-7c], ebx
'0071f6bb    899d74ffffff            mov dword ptr [ebp+ffffff74], ebx
'0071f6c1    899d64ffffff            mov dword ptr [ebp+ffffff64], ebx
'0071f6c7    899d30ffffff            mov dword ptr [ebp+ffffff30], ebx
'0071f6cd    899d2cffffff            mov dword ptr [ebp+ffffff2c], ebx
'0071f6d3    ff9210030000            call dword ptr [edx+00000310]
'0071f6d9    8945ac                  mov dword ptr [ebp-54], eax
'0071f6dc    8d45a4                  lea eax, dword ptr [ebp-5c]
'0071f6df    50                      push eax
'0071f6e0    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0071f6e3    51                      push ecx
'0071f6e4    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcTrimVar"
'0071f6eb    ff15e4104000            call dword ptr [004010e4]
'0071f6f1    8d5594                  lea edx, dword ptr [ebp-6c]
'0071f6f4    52                      push edx
'0071f6f5    8d8564ffffff            lea eax, dword ptr [ebp+ffffff64]
'0071f6fb    50                      push eax
'0071f6fc    c7856cffffffcc134300    mov dword ptr [ebp+ffffff6c], 004313cc
'0071f706    c78564ffffff08800000    mov dword ptr [ebp+ffffff64], 00008008

' *** Reference to "__vbaVarTstEq"
'0071f710    ff153c114000            call dword ptr [0040113c]
'0071f716    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0071f719    51                      push ecx
'0071f71a    8d55a4                  lea edx, dword ptr [ebp-5c]
'0071f71d    52                      push edx
'0071f71e    6a02                    push 02
'0071f720    668bf8                  mov di, ax

' *** Reference to "__vbaFreeVarList"
'0071f723    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_121)
'0071f729    83c40c                  add esp, 0c
'0071f72c    663bfb                  cmp di, bx
'0071f72f    0f84ab000000            je 71f7e0

If (((Trim(Me)) = (vbNullChar))) Then
'0071f735    b904000280              mov ecx, 80020004
'0071f73a    b80a000000              mov eax, 0000000a
'0071f73f    898d7cffffff            mov dword ptr [ebp+ffffff7c], ecx
'0071f745    894d8c                  mov dword ptr [ebp-74], ecx
'0071f748    894d9c                  mov dword ptr [ebp-64], ecx
'0071f74b    8d9564ffffff            lea edx, dword ptr [ebp+ffffff64]
'0071f751    8d4da4                  lea ecx, dword ptr [ebp-5c]
'0071f754    898574ffffff            mov dword ptr [ebp+ffffff74], eax
'0071f75a    894584                  mov dword ptr [ebp-7c], eax
'0071f75d    894594                  mov dword ptr [ebp-6c], eax
'0071f760    c7856cffffff50704500    mov dword ptr [ebp+ffffff6c], 00457050
'0071f76a    c78564ffffff08000000    mov dword ptr [ebp+ffffff64], 00000008

' *** Reference to "__vbaVarDup"
'0071f774    ff1598124000            call dword ptr [00401298]
    var_83 = ("Saisie d'un nom obligatoire")
'0071f77a    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'0071f780    50                      push eax
'0071f781    8d4d84                  lea ecx, dword ptr [ebp-7c]
'0071f784    51                      push ecx
'0071f785    8d5594                  lea edx, dword ptr [ebp-6c]
'0071f788    52                      push edx
'0071f789    53                      push ebx
'0071f78a    8d45a4                  lea eax, dword ptr [ebp-5c]
'0071f78d    50                      push eax

' *** Reference to "rtcMsgBox"
'0071f78e    ff15b8104000            call dword ptr [004010b8]
    var_76 = MsgBox(var_83, 0)
'0071f794    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'0071f79a    51                      push ecx
'0071f79b    8d5584                  lea edx, dword ptr [ebp-7c]
'0071f79e    52                      push edx
'0071f79f    8d4594                  lea eax, dword ptr [ebp-6c]
'0071f7a2    50                      push eax
'0071f7a3    8d4da4                  lea ecx, dword ptr [ebp-5c]
'0071f7a6    51                      push ecx
'0071f7a7    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0071f7a9    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_83, var_121, var_119, var_91)
'0071f7af    8b16                    mov edx, dword ptr [esi]
'0071f7b1    83c414                  add esp, 14
'0071f7b4    56                      push esi
'0071f7b5    ff9210030000            call dword ptr [edx+00000310]
'0071f7bb    50                      push eax
'0071f7bc    8d45c0                  lea eax, dword ptr [ebp-40]
'0071f7bf    50                      push eax

' *** Reference to "__vbaObjSet"
'0071f7c0    ff15b4104000            call dword ptr [004010b4]
    Set var_5 = 
'0071f7c6    8bf0                    mov esi, eax
'0071f7c8    8b0e                    mov ecx, dword ptr [esi]
'0071f7ca    56                      push esi
'0071f7cb    ff9104020000            call dword ptr [ecx+00000204]
'0071f7d1    dbe2                    fnclex
'0071f7d3    3bc3                    cmp eax, ebx
'0071f7d5    0f8dc3020000            jge 71fa9e
    
    If (    var_5 < 0) Then
'0071f7db    e9ac020000              jmp 71fa8c
    
End If
'0071f7e0    8b16                    mov edx, dword ptr [esi]
'0071f7e2    56                      push esi
'0071f7e3    ff9218030000            call dword ptr [edx+00000318]
'0071f7e9    8945ac                  mov dword ptr [ebp-54], eax
'0071f7ec    8d45a4                  lea eax, dword ptr [ebp-5c]
'0071f7ef    50                      push eax
'0071f7f0    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNumeric"
'0071f7f7    ff154c114000            call dword ptr [0040114c]
'0071f7fd    33ff                    xor edi, edi
var_num7 = Empty
'0071f7ff    668bf8                  mov di, ax
'0071f802    66f7df                  neg di
'0071f805    8d4da4                  lea ecx, dword ptr [ebp-5c]
'0071f808    1bff                    sbb edi, edi
'0071f80a    47                      inc edi
'0071f80b    f7df                    neg edi

' *** Reference to "__vbaFreeVar"
'0071f80d    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_83)
'0071f813    663bfb                  cmp di, bx
'0071f816    0f84ab000000            je 71f8c7

If (IsNumeric(((Trim(Me)) [#-@-@#] (vbNullChar)))) Then
'0071f81c    b904000280              mov ecx, 80020004
'0071f821    b80a000000              mov eax, 0000000a
'0071f826    898d7cffffff            mov dword ptr [ebp+ffffff7c], ecx
'0071f82c    894d8c                  mov dword ptr [ebp-74], ecx
'0071f82f    894d9c                  mov dword ptr [ebp-64], ecx
'0071f832    8d9564ffffff            lea edx, dword ptr [ebp+ffffff64]
'0071f838    8d4da4                  lea ecx, dword ptr [ebp-5c]
'0071f83b    898574ffffff            mov dword ptr [ebp+ffffff74], eax
'0071f841    894584                  mov dword ptr [ebp-7c], eax
'0071f844    894594                  mov dword ptr [ebp-6c], eax
'0071f847    c7856cffffffcc704500    mov dword ptr [ebp+ffffff6c], 004570cc
'0071f851    c78564ffffff08000000    mov dword ptr [ebp+ffffff64], 00000008

' *** Reference to "__vbaVarDup"
'0071f85b    ff1598124000            call dword ptr [00401298]
    var_83 = ("Saisie d'une valeur numérique pour la population")
'0071f861    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'0071f867    51                      push ecx
'0071f868    8d5584                  lea edx, dword ptr [ebp-7c]
'0071f86b    52                      push edx
'0071f86c    8d4594                  lea eax, dword ptr [ebp-6c]
'0071f86f    50                      push eax
'0071f870    53                      push ebx
'0071f871    8d4da4                  lea ecx, dword ptr [ebp-5c]
'0071f874    51                      push ecx

' *** Reference to "rtcMsgBox"
'0071f875    ff15b8104000            call dword ptr [004010b8]
    var_127 = MsgBox(var_83, 0)
'0071f87b    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'0071f881    52                      push edx
'0071f882    8d4584                  lea eax, dword ptr [ebp-7c]
'0071f885    50                      push eax
'0071f886    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0071f889    51                      push ecx
'0071f88a    8d55a4                  lea edx, dword ptr [ebp-5c]
'0071f88d    52                      push edx
'0071f88e    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0071f890    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_83, var_121, var_119, var_91)
'0071f896    8b06                    mov eax, dword ptr [esi]
'0071f898    83c414                  add esp, 14
'0071f89b    56                      push esi
'0071f89c    ff9018030000            call dword ptr [eax+00000318]
'0071f8a2    50                      push eax
'0071f8a3    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0071f8a6    51                      push ecx

' *** Reference to "__vbaObjSet"
'0071f8a7    ff15b4104000            call dword ptr [004010b4]
    Set var_5 = Nothing
'0071f8ad    8bf0                    mov esi, eax
'0071f8af    8b16                    mov edx, dword ptr [esi]
'0071f8b1    56                      push esi
'0071f8b2    ff9204020000            call dword ptr [edx+00000204]
'0071f8b8    dbe2                    fnclex
'0071f8ba    3bc3                    cmp eax, ebx
'0071f8bc    0f8ddc010000            jge 71fa9e
    
    If (    var_5 < 0) Then
'0071f8c2    e9c5010000              jmp 71fa8c
    
End If
'0071f8c7    8b06                    mov eax, dword ptr [esi]
'0071f8c9    56                      push esi
'0071f8ca    ff900c030000            call dword ptr [eax+0000030c]
'0071f8d0    8d4da4                  lea ecx, dword ptr [ebp-5c]
'0071f8d3    51                      push ecx
'0071f8d4    8945ac                  mov dword ptr [ebp-54], eax
'0071f8d7    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNumeric"
'0071f8de    ff154c114000            call dword ptr [0040114c]
'0071f8e4    33ff                    xor edi, edi
var_num7 = Empty
'0071f8e6    668bf8                  mov di, ax
'0071f8e9    66f7df                  neg di
'0071f8ec    8d4da4                  lea ecx, dword ptr [ebp-5c]
'0071f8ef    1bff                    sbb edi, edi
'0071f8f1    47                      inc edi
'0071f8f2    f7df                    neg edi

' *** Reference to "__vbaFreeVar"
'0071f8f4    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_83)
'0071f8fa    663bfb                  cmp di, bx
'0071f8fd    0f84ab000000            je 71f9ae

If (-(CBool((IsNumeric(frmGenerateurVille)))) <> 0) Then
'0071f903    b904000280              mov ecx, 80020004
'0071f908    b80a000000              mov eax, 0000000a
'0071f90d    898d7cffffff            mov dword ptr [ebp+ffffff7c], ecx
'0071f913    894d8c                  mov dword ptr [ebp-74], ecx
'0071f916    894d9c                  mov dword ptr [ebp-64], ecx
'0071f919    8d9564ffffff            lea edx, dword ptr [ebp+ffffff64]
'0071f91f    8d4da4                  lea ecx, dword ptr [ebp-5c]
'0071f922    898574ffffff            mov dword ptr [ebp+ffffff74], eax
'0071f928    894584                  mov dword ptr [ebp-7c], eax
'0071f92b    894594                  mov dword ptr [ebp-6c], eax
'0071f92e    c7856cffffff34714500    mov dword ptr [ebp+ffffff6c], 00457134
'0071f938    c78564ffffff08000000    mov dword ptr [ebp+ffffff64], 00000008

' *** Reference to "__vbaVarDup"
'0071f942    ff1598124000            call dword ptr [00401298]
    var_83 = ("Saisie d'une valeur numérique pour les impôts")
'0071f948    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'0071f94e    52                      push edx
'0071f94f    8d4584                  lea eax, dword ptr [ebp-7c]
'0071f952    50                      push eax
'0071f953    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0071f956    51                      push ecx
'0071f957    53                      push ebx
'0071f958    8d55a4                  lea edx, dword ptr [ebp-5c]
'0071f95b    52                      push edx

' *** Reference to "rtcMsgBox"
'0071f95c    ff15b8104000            call dword ptr [004010b8]
    var_17 = MsgBox(var_83, 0)
'0071f962    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'0071f968    50                      push eax
'0071f969    8d4d84                  lea ecx, dword ptr [ebp-7c]
'0071f96c    51                      push ecx
'0071f96d    8d5594                  lea edx, dword ptr [ebp-6c]
'0071f970    52                      push edx
'0071f971    8d45a4                  lea eax, dword ptr [ebp-5c]
'0071f974    50                      push eax
'0071f975    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0071f977    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_83, var_121, var_119, var_91)
'0071f97d    8b0e                    mov ecx, dword ptr [esi]
'0071f97f    83c414                  add esp, 14
'0071f982    56                      push esi
'0071f983    ff910c030000            call dword ptr [ecx+0000030c]
'0071f989    50                      push eax
'0071f98a    8d55c0                  lea edx, dword ptr [ebp-40]
'0071f98d    52                      push edx

' *** Reference to "__vbaObjSet"
'0071f98e    ff15b4104000            call dword ptr [004010b4]
    Set var_5 = 
'0071f994    8bf0                    mov esi, eax
'0071f996    8b06                    mov eax, dword ptr [esi]
'0071f998    56                      push esi
'0071f999    ff9004020000            call dword ptr [eax+00000204]
'0071f99f    dbe2                    fnclex
'0071f9a1    3bc3                    cmp eax, ebx
'0071f9a3    0f8df5000000            jge 71fa9e
    
    If (    var_83 < 0) Then
'0071f9a9    e9de000000              jmp 71fa8c
    
End If
'0071f9ae    8b0e                    mov ecx, dword ptr [esi]
'0071f9b0    56                      push esi
'0071f9b1    ff9108030000            call dword ptr [ecx+00000308]
'0071f9b7    8d55a4                  lea edx, dword ptr [ebp-5c]
'0071f9ba    52                      push edx
'0071f9bb    8945ac                  mov dword ptr [ebp-54], eax
'0071f9be    c745a409000000          mov dword ptr [ebp-5c], 00000009

' *** Reference to "rtcIsNumeric"
'0071f9c5    ff154c114000            call dword ptr [0040114c]
'0071f9cb    33ff                    xor edi, edi
var_num7 = Empty
'0071f9cd    668bf8                  mov di, ax
'0071f9d0    66f7df                  neg di
'0071f9d3    8d4da4                  lea ecx, dword ptr [ebp-5c]
'0071f9d6    1bff                    sbb edi, edi
'0071f9d8    47                      inc edi
'0071f9d9    f7df                    neg edi

' *** Reference to "__vbaFreeVar"
'0071f9db    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_83)
'0071f9e1    663bfb                  cmp di, bx
'0071f9e4    0f84c2000000            je 71faac

If (-(CBool((IsNumeric(IsNumeric(frmGenerateurVille))))) <> 0) Then
'0071f9ea    b904000280              mov ecx, 80020004
'0071f9ef    b80a000000              mov eax, 0000000a
'0071f9f4    898d7cffffff            mov dword ptr [ebp+ffffff7c], ecx
'0071f9fa    894d8c                  mov dword ptr [ebp-74], ecx
'0071f9fd    894d9c                  mov dword ptr [ebp-64], ecx
'0071fa00    8d9564ffffff            lea edx, dword ptr [ebp+ffffff64]
'0071fa06    8d4da4                  lea ecx, dword ptr [ebp-5c]
'0071fa09    898574ffffff            mov dword ptr [ebp+ffffff74], eax
'0071fa0f    894584                  mov dword ptr [ebp-7c], eax
'0071fa12    894594                  mov dword ptr [ebp-6c], eax
'0071fa15    c7856cffffff94714500    mov dword ptr [ebp+ffffff6c], 00457194
'0071fa1f    c78564ffffff08000000    mov dword ptr [ebp+ffffff64], 00000008

' *** Reference to "__vbaVarDup"
'0071fa29    ff1598124000            call dword ptr [00401298]
    var_83 = ("Saisie d'une valeur numérique pour la dîme")
'0071fa2f    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'0071fa35    50                      push eax
'0071fa36    8d4d84                  lea ecx, dword ptr [ebp-7c]
'0071fa39    51                      push ecx
'0071fa3a    8d5594                  lea edx, dword ptr [ebp-6c]
'0071fa3d    52                      push edx
'0071fa3e    53                      push ebx
'0071fa3f    8d45a4                  lea eax, dword ptr [ebp-5c]
'0071fa42    50                      push eax

' *** Reference to "rtcMsgBox"
'0071fa43    ff15b8104000            call dword ptr [004010b8]
    var_54 = MsgBox(var_83, 0)
'0071fa49    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'0071fa4f    51                      push ecx
'0071fa50    8d5584                  lea edx, dword ptr [ebp-7c]
'0071fa53    52                      push edx
'0071fa54    8d4594                  lea eax, dword ptr [ebp-6c]
'0071fa57    50                      push eax
'0071fa58    8d4da4                  lea ecx, dword ptr [ebp-5c]
'0071fa5b    51                      push ecx
'0071fa5c    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0071fa5e    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_83, var_121, var_119, var_91)
'0071fa64    8b16                    mov edx, dword ptr [esi]
'0071fa66    83c414                  add esp, 14
'0071fa69    56                      push esi
'0071fa6a    ff9208030000            call dword ptr [edx+00000308]
'0071fa70    50                      push eax
'0071fa71    8d45c0                  lea eax, dword ptr [ebp-40]
'0071fa74    50                      push eax

' *** Reference to "__vbaObjSet"
'0071fa75    ff15b4104000            call dword ptr [004010b4]
    Set var_5 = 
'0071fa7b    8bf0                    mov esi, eax
'0071fa7d    8b0e                    mov ecx, dword ptr [esi]
'0071fa7f    56                      push esi
'0071fa80    ff9104020000            call dword ptr [ecx+00000204]
'0071fa86    dbe2                    fnclex
'0071fa88    3bc3                    cmp eax, ebx
'0071fa8a    7d12                    jge 71fa9e
    
    If (    var_5 < 0) Then
'0071fa8c    6804020000              push 00000204
'0071fa91    68d00d4300              push 00430dd0
'0071fa96    56                      push esi
'0071fa97    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071fa98    ff1580104000            call dword ptr [00401080]
    
End If
'0071fa9e    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeObj"
'0071faa1    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_5)
'0071faa7    e9b5200000              jmp 721b61

'ERROR: Two many next close:
End If
'0071faac    8b16                    mov edx, dword ptr [esi]
'0071faae    56                      push esi
'0071faaf    ff9204030000            call dword ptr [edx+00000304]
'0071fab5    50                      push eax
'0071fab6    8d45c0                  lea eax, dword ptr [ebp-40]
'0071fab9    50                      push eax

' *** Reference to "__vbaObjSet"
'0071faba    ff15b4104000            call dword ptr [004010b4]
Set var_5 = IsNumeric(IsNumeric(frmGenerateurVille))
'0071fac0    8bf8                    mov edi, eax
'0071fac2    8b0f                    mov ecx, dword ptr [edi]
'0071fac4    68cc134300              push 004313cc
'0071fac9    57                      push edi
'0071faca    ff91a4000000            call dword ptr [ecx+000000a4]
'0071fad0    dbe2                    fnclex
'0071fad2    3bc3                    cmp eax, ebx
'0071fad4    7d12                    jge 71fae8

If (var_5 < 0) Then
'0071fad6    68a4000000              push 000000a4
'0071fadb    68d00d4300              push 00430dd0
'0071fae0    57                      push edi
'0071fae1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071fae2    ff1580104000            call dword ptr [00401080]
    
End If
'0071fae8    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeObj"
'0071faeb    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_5)
'0071faf1    8b16                    mov edx, dword ptr [esi]
'0071faf3    56                      push esi
'0071faf4    ff9204030000            call dword ptr [edx+00000304]

' *** Reference to "__vbaObjSet"
'0071fafa    8b3db4104000            mov edi, dword ptr [004010b4]
'0071fb00    50                      push eax
'0071fb01    8d45b4                  lea eax, dword ptr [ebp-4c]
'0071fb04    50                      push eax
'0071fb05    ffd7                    call edi
Set var_62 = var_5
'0071fb07    8b0e                    mov ecx, dword ptr [esi]
'0071fb09    56                      push esi
'0071fb0a    89850cffffff            mov dword ptr [ebp+ffffff0c], eax
'0071fb10    ff9110030000            call dword ptr [ecx+00000310]
'0071fb16    50                      push eax
'0071fb17    8d55c0                  lea edx, dword ptr [ebp-40]
'0071fb1a    52                      push edx
'0071fb1b    ffd7                    call edi
Set var_5 = var_62
'0071fb1d    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071fb20    8bf8                    mov edi, eax
'0071fb22    8b07                    mov eax, dword ptr [edi]
'0071fb24    51                      push ecx
'0071fb25    57                      push edi
'0071fb26    ff90a0000000            call dword ptr [eax+000000a0]
'0071fb2c    dbe2                    fnclex
'0071fb2e    3bc3                    cmp eax, ebx
'0071fb30    7d12                    jge 71fb44

If (0 < 0) Then
'0071fb32    68a0000000              push 000000a0
'0071fb37    68d00d4300              push 00430dd0
'0071fb3c    57                      push edi
'0071fb3d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071fb3e    ff1580104000            call dword ptr [00401080]
    
End If
'0071fb44    8b16                    mov edx, dword ptr [esi]
'0071fb46    56                      push esi
'0071fb47    ff9218030000            call dword ptr [edx+00000318]
'0071fb4d    50                      push eax
'0071fb4e    8d45bc                  lea eax, dword ptr [ebp-44]
'0071fb51    50                      push eax

' *** Reference to "__vbaObjSet"
'0071fb52    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Nothing
'0071fb58    8d55e4                  lea edx, dword ptr [ebp-1c]
'0071fb5b    8bf8                    mov edi, eax
'0071fb5d    8b0f                    mov ecx, dword ptr [edi]
'0071fb5f    52                      push edx
'0071fb60    57                      push edi
'0071fb61    ff91a0000000            call dword ptr [ecx+000000a0]
'0071fb67    dbe2                    fnclex
'0071fb69    3bc3                    cmp eax, ebx
'0071fb6b    7d12                    jge 71fb7f

If (var_58 < 0) Then
'0071fb6d    68a0000000              push 000000a0
'0071fb72    68d00d4300              push 00430dd0
'0071fb77    57                      push edi
'0071fb78    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071fb79    ff1580104000            call dword ptr [00401080]
    
End If
'0071fb7f    8b45e4                  mov eax, dword ptr [ebp-1c]
'0071fb82    50                      push eax

' *** Reference to "__vbaI4Str"
'0071fb83    ff1550124000            call dword ptr [00401250]
'0071fb89    8b0e                    mov ecx, dword ptr [esi]
'0071fb8b    8d55e0                  lea edx, dword ptr [ebp-20]
'0071fb8e    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'0071fb94    52                      push edx
'0071fb95    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'0071fb9b    50                      push eax
'0071fb9c    56                      push esi
'0071fb9d    ff91f8060000            call dword ptr [ecx+000006f8]
'0071fba3    3bc3                    cmp eax, ebx
'0071fba5    7d12                    jge 71fbb9
'0071fba7    68f8060000              push 000006f8
'0071fbac    68b0154300              push 004315b0
'0071fbb1    56                      push esi
'0071fbb2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071fbb3    ff1580104000            call dword ptr [00401080]
'0071fbb9    8b0e                    mov ecx, dword ptr [esi]
'0071fbbb    56                      push esi
'0071fbbc    ff9118030000            call dword ptr [ecx+00000318]
'0071fbc2    50                      push eax
'0071fbc3    8d55b8                  lea edx, dword ptr [ebp-48]
'0071fbc6    52                      push edx

' *** Reference to "__vbaObjSet"
'0071fbc7    ff15b4104000            call dword ptr [004010b4]
Set var_61 = 
'0071fbcd    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0071fbd0    8bf8                    mov edi, eax
'0071fbd2    8b07                    mov eax, dword ptr [edi]
'0071fbd4    51                      push ecx
'0071fbd5    57                      push edi
'0071fbd6    ff90a0000000            call dword ptr [eax+000000a0]
'0071fbdc    dbe2                    fnclex
'0071fbde    3bc3                    cmp eax, ebx
'0071fbe0    7d12                    jge 71fbf4

If (CLng(vbNullString) < 0) Then
'0071fbe2    68a0000000              push 000000a0
'0071fbe7    68d00d4300              push 00430dd0
'0071fbec    57                      push edi
'0071fbed    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071fbee    ff1580104000            call dword ptr [00401080]
    
End If
'0071fbf4    8b45e8                  mov eax, dword ptr [ebp-18]
'0071fbf7    8b950cffffff            mov edx, dword ptr [ebp+ffffff0c]
'0071fbfd    8b1a                    mov ebx, dword ptr [edx]
'0071fbff    50                      push eax
'0071fc00    68fc8c4300              push 00438cfc

' *** Reference to "__vbaStrCat"
'0071fc05    ff1570104000            call dword ptr [00401070]
var_49 = (vbNullString) & (", ")

' *** Reference to "__vbaStrMove"
'0071fc0b    8b3dd0124000            mov edi, dword ptr [004012d0]
'0071fc11    8bd0                    mov edx, eax
'0071fc13    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0071fc16    ffd7                    call edi
'0071fc18    8b4de0                  mov ecx, dword ptr [ebp-20]
'0071fc1b    50                      push eax
'0071fc1c    51                      push ecx

' *** Reference to "__vbaStrCat"
'0071fc1d    ff1570104000            call dword ptr [00401070]
var_56 = (var_49) & (vbNullString)
'0071fc23    8bd0                    mov edx, eax
'0071fc25    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0071fc28    ffd7                    call edi
'0071fc2a    50                      push eax
'0071fc2b    68903c4400              push 00443c90

' *** Reference to "__vbaStrCat"
'0071fc30    ff1570104000            call dword ptr [00401070]
var_2462 = (var_56) & (" de ")
'0071fc36    8bd0                    mov edx, eax
'0071fc38    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0071fc3b    ffd7                    call edi
'0071fc3d    8b55d4                  mov edx, dword ptr [ebp-2c]
'0071fc40    50                      push eax
'0071fc41    52                      push edx

' *** Reference to "__vbaStrCat"
'0071fc42    ff1570104000            call dword ptr [00401070]
var_286 = (var_2462) & (vbNullString)
'0071fc48    8bd0                    mov edx, eax
'0071fc4a    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0071fc4d    ffd7                    call edi
'0071fc4f    50                      push eax
'0071fc50    68f0714500              push 004571f0

' *** Reference to "__vbaStrCat"
'0071fc55    ff1570104000            call dword ptr [00401070]
var_1231 = (var_286) & (" habitants.")
'0071fc5b    8bd0                    mov edx, eax
'0071fc5d    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0071fc60    ffd7                    call edi
'0071fc62    50                      push eax
'0071fc63    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'0071fc68    ff1570104000            call dword ptr [00401070]
var_2068 = (var_1231) & (vbCrLf)
'0071fc6e    8bd0                    mov edx, eax
'0071fc70    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0071fc73    ffd7                    call edi
'0071fc75    50                      push eax
'0071fc76    8bc3                    mov eax, ebx
'0071fc78    8b9d0cffffff            mov ebx, dword ptr [ebp+ffffff0c]
'0071fc7e    53                      push ebx
'0071fc7f    ff90a4000000            call dword ptr [eax+000000a4]
'0071fc85    dbe2                    fnclex
'0071fc87    85c0                    test ax, ax
'0071fc89    7d12                    jge 71fc9d
'0071fc8b    68a4000000              push 000000a4
'0071fc90    68d00d4300              push 00430dd0
'0071fc95    53                      push ebx
'0071fc96    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071fc97    ff1580104000            call dword ptr [00401080]
'0071fc9d    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0071fca0    51                      push ecx
'0071fca1    8d55c8                  lea edx, dword ptr [ebp-38]
'0071fca4    52                      push edx
'0071fca5    8d45cc                  lea eax, dword ptr [ebp-34]
'0071fca8    50                      push eax
'0071fca9    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0071fcac    51                      push ecx
'0071fcad    8d55d0                  lea edx, dword ptr [ebp-30]
'0071fcb0    52                      push edx
'0071fcb1    8d45d8                  lea eax, dword ptr [ebp-28]
'0071fcb4    50                      push eax
'0071fcb5    8d4de0                  lea ecx, dword ptr [ebp-20]
'0071fcb8    51                      push ecx
'0071fcb9    8d55dc                  lea edx, dword ptr [ebp-24]
'0071fcbc    52                      push edx
'0071fcbd    8d45e4                  lea eax, dword ptr [ebp-1c]
'0071fcc0    50                      push eax
'0071fcc1    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071fcc4    51                      push ecx
'0071fcc5    6a0a                    push 0a

' *** Reference to "__vbaFreeStrList"
'0071fcc7    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, 0, -4568, 0, -4572, -4576, 0, -4580, -4584, -4588)
'0071fccd    8d55b4                  lea edx, dword ptr [ebp-4c]
'0071fcd0    52                      push edx
'0071fcd1    8d45b8                  lea eax, dword ptr [ebp-48]
'0071fcd4    50                      push eax
'0071fcd5    8d4dbc                  lea ecx, dword ptr [ebp-44]
'0071fcd8    51                      push ecx
'0071fcd9    8d55c0                  lea edx, dword ptr [ebp-40]
'0071fcdc    52                      push edx
'0071fcdd    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'0071fcdf    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58, var_61, var_62)
'0071fce5    8b06                    mov eax, dword ptr [esi]
'0071fce7    83c440                  add esp, 40
'0071fcea    56                      push esi
'0071fceb    ff9004030000            call dword ptr [eax+00000304]

' *** Reference to "__vbaObjSet"
'0071fcf1    8b1db4104000            mov ebx, dword ptr [004010b4]
'0071fcf7    50                      push eax
'0071fcf8    8d4db8                  lea ecx, dword ptr [ebp-48]
'0071fcfb    51                      push ecx
'0071fcfc    ffd3                    call ebx
Set var_61 = Nothing
'0071fcfe    8b16                    mov edx, dword ptr [esi]
'0071fd00    56                      push esi
'0071fd01    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'0071fd07    ff9204030000            call dword ptr [edx+00000304]
'0071fd0d    50                      push eax
'0071fd0e    8d45c0                  lea eax, dword ptr [ebp-40]
'0071fd11    50                      push eax
'0071fd12    ffd3                    call ebx
Set var_5 = Nothing
'0071fd14    8d55e8                  lea edx, dword ptr [ebp-18]
'0071fd17    8bd8                    mov ebx, eax
'0071fd19    8b0b                    mov ecx, dword ptr [ebx]
'0071fd1b    52                      push edx
'0071fd1c    53                      push ebx
'0071fd1d    ff91a0000000            call dword ptr [ecx+000000a0]
'0071fd23    dbe2                    fnclex
'0071fd25    85c0                    test ax, ax
'0071fd27    7d12                    jge 71fd3b
'0071fd29    68a0000000              push 000000a0
'0071fd2e    68d00d4300              push 00430dd0
'0071fd33    53                      push ebx
'0071fd34    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071fd35    ff1580104000            call dword ptr [00401080]
'0071fd3b    8b06                    mov eax, dword ptr [esi]
'0071fd3d    56                      push esi
'0071fd3e    ff9018030000            call dword ptr [eax+00000318]
'0071fd44    50                      push eax
'0071fd45    8d4dbc                  lea ecx, dword ptr [ebp-44]
'0071fd48    51                      push ecx

' *** Reference to "__vbaObjSet"
'0071fd49    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Nothing
'0071fd4f    8bd8                    mov ebx, eax
'0071fd51    8b13                    mov edx, dword ptr [ebx]
'0071fd53    8d45e4                  lea eax, dword ptr [ebp-1c]
'0071fd56    50                      push eax
'0071fd57    53                      push ebx
'0071fd58    ff92a0000000            call dword ptr [edx+000000a0]
'0071fd5e    dbe2                    fnclex
'0071fd60    85c0                    test ax, ax
'0071fd62    7d12                    jge 71fd76
'0071fd64    68a0000000              push 000000a0
'0071fd69    68d00d4300              push 00430dd0
'0071fd6e    53                      push ebx
'0071fd6f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071fd70    ff1580104000            call dword ptr [00401080]
'0071fd76    8b4de4                  mov ecx, dword ptr [ebp-1c]
'0071fd79    51                      push ecx

' *** Reference to "__vbaI4Str"
'0071fd7a    ff1550124000            call dword ptr [00401250]
'0071fd80    8b16                    mov edx, dword ptr [esi]
'0071fd82    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'0071fd88    8d852cffffff            lea eax, dword ptr [ebp+ffffff2c]
'0071fd8e    50                      push eax
'0071fd8f    8d8d30ffffff            lea ecx, dword ptr [ebp+ffffff30]
'0071fd95    51                      push ecx
'0071fd96    56                      push esi
'0071fd97    ff92fc060000            call dword ptr [edx+000006fc]
'0071fd9d    85c0                    test ax, ax
'0071fd9f    7d12                    jge 71fdb3
'0071fda1    68fc060000              push 000006fc
'0071fda6    68b0154300              push 004315b0
'0071fdab    56                      push esi
'0071fdac    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071fdad    ff1580104000            call dword ptr [00401080]
'0071fdb3    8b45e8                  mov eax, dword ptr [ebp-18]
'0071fdb6    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'0071fdbc    8b1a                    mov ebx, dword ptr [edx]
'0071fdbe    50                      push eax
'0071fdbf    680c724500              push 0045720c

' *** Reference to "__vbaStrCat"
'0071fdc4    ff1570104000            call dword ptr [00401070]
var_49 = (vbNullString) & ("Limite financière : ")
'0071fdca    8bd0                    mov edx, eax
'0071fdcc    8d4de0                  lea ecx, dword ptr [ebp-20]
'0071fdcf    ffd7                    call edi
'0071fdd1    8b8d2cffffff            mov ecx, dword ptr [ebp+ffffff2c]
'0071fdd7    50                      push eax
'0071fdd8    51                      push ecx

' *** Reference to "__vbaStrI4"
'0071fdd9    ff1520104000            call dword ptr [00401020]
'0071fddf    8bd0                    mov edx, eax
'0071fde1    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0071fde4    ffd7                    call edi
'0071fde6    50                      push eax

' *** Reference to "__vbaStrCat"
'0071fde7    ff1570104000            call dword ptr [00401070]
var_868 = (var_49) & (CStr(0))
'0071fded    8bd0                    mov edx, eax
'0071fdef    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0071fdf2    ffd7                    call edi
'0071fdf4    50                      push eax
'0071fdf5    683c724500              push 0045723c

' *** Reference to "__vbaStrCat"
'0071fdfa    ff1570104000            call dword ptr [00401070]
var_2461 = (var_868) & (" Po")
'0071fe00    8bd0                    mov edx, eax
'0071fe02    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0071fe05    ffd7                    call edi
'0071fe07    50                      push eax
'0071fe08    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'0071fe0d    ff1570104000            call dword ptr [00401070]
var_870 = (var_2461) & (vbCrLf)
'0071fe13    8bd0                    mov edx, eax
'0071fe15    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0071fe18    ffd7                    call edi
'0071fe1a    8bd3                    mov edx, ebx
'0071fe1c    8b9d14ffffff            mov ebx, dword ptr [ebp+ffffff14]
'0071fe22    50                      push eax
'0071fe23    53                      push ebx
'0071fe24    ff92a4000000            call dword ptr [edx+000000a4]
'0071fe2a    dbe2                    fnclex
'0071fe2c    85c0                    test ax, ax
'0071fe2e    7d12                    jge 71fe42
'0071fe30    68a4000000              push 000000a4
'0071fe35    68d00d4300              push 00430dd0
'0071fe3a    53                      push ebx
'0071fe3b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071fe3c    ff1580104000            call dword ptr [00401080]
'0071fe42    8d45d0                  lea eax, dword ptr [ebp-30]
'0071fe45    50                      push eax
'0071fe46    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0071fe49    51                      push ecx
'0071fe4a    8d55d8                  lea edx, dword ptr [ebp-28]
'0071fe4d    52                      push edx
'0071fe4e    8d45dc                  lea eax, dword ptr [ebp-24]
'0071fe51    50                      push eax
'0071fe52    8d4de0                  lea ecx, dword ptr [ebp-20]
'0071fe55    51                      push ecx
'0071fe56    8d55e4                  lea edx, dword ptr [ebp-1c]
'0071fe59    52                      push edx
'0071fe5a    8d45e8                  lea eax, dword ptr [ebp-18]
'0071fe5d    50                      push eax
'0071fe5e    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0071fe60    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, 0, -4596, -4600, -4604, -4608, -4612)
'0071fe66    8d4db8                  lea ecx, dword ptr [ebp-48]
'0071fe69    51                      push ecx
'0071fe6a    8d55bc                  lea edx, dword ptr [ebp-44]
'0071fe6d    52                      push edx
'0071fe6e    8d45c0                  lea eax, dword ptr [ebp-40]
'0071fe71    50                      push eax
'0071fe72    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0071fe74    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58, var_61)
'0071fe7a    8b0e                    mov ecx, dword ptr [esi]
'0071fe7c    83c430                  add esp, 30
'0071fe7f    56                      push esi
'0071fe80    ff9104030000            call dword ptr [ecx+00000304]

' *** Reference to "__vbaObjSet"
'0071fe86    8b1db4104000            mov ebx, dword ptr [004010b4]
'0071fe8c    50                      push eax
'0071fe8d    8d55b8                  lea edx, dword ptr [ebp-48]
'0071fe90    52                      push edx
'0071fe91    ffd3                    call ebx
Set var_61 = 
'0071fe93    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'0071fe99    8b06                    mov eax, dword ptr [esi]
'0071fe9b    56                      push esi
'0071fe9c    ff9004030000            call dword ptr [eax+00000304]
'0071fea2    50                      push eax
'0071fea3    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0071fea6    51                      push ecx
'0071fea7    ffd3                    call ebx
Set var_5 = Nothing
'0071fea9    8bd8                    mov ebx, eax
'0071feab    8b13                    mov edx, dword ptr [ebx]
'0071fead    8d45e8                  lea eax, dword ptr [ebp-18]
'0071feb0    50                      push eax
'0071feb1    53                      push ebx
'0071feb2    ff92a0000000            call dword ptr [edx+000000a0]
'0071feb8    dbe2                    fnclex
'0071feba    85c0                    test ax, ax
'0071febc    7d12                    jge 71fed0
'0071febe    68a0000000              push 000000a0
'0071fec3    68d00d4300              push 00430dd0
'0071fec8    53                      push ebx
'0071fec9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071feca    ff1580104000            call dword ptr [00401080]
'0071fed0    8b0e                    mov ecx, dword ptr [esi]
'0071fed2    56                      push esi
'0071fed3    ff9118030000            call dword ptr [ecx+00000318]
'0071fed9    50                      push eax
'0071feda    8d55bc                  lea edx, dword ptr [ebp-44]
'0071fedd    52                      push edx

' *** Reference to "__vbaObjSet"
'0071fede    ff15b4104000            call dword ptr [004010b4]
Set var_58 = 
'0071fee4    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0071fee7    8bd8                    mov ebx, eax
'0071fee9    8b03                    mov eax, dword ptr [ebx]
'0071feeb    51                      push ecx
'0071feec    53                      push ebx
'0071feed    ff90a0000000            call dword ptr [eax+000000a0]
'0071fef3    dbe2                    fnclex
'0071fef5    85c0                    test ax, ax
'0071fef7    7d12                    jge 71ff0b
'0071fef9    68a0000000              push 000000a0
'0071fefe    68d00d4300              push 00430dd0
'0071ff03    53                      push ebx
'0071ff04    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071ff05    ff1580104000            call dword ptr [00401080]
'0071ff0b    8b55e4                  mov edx, dword ptr [ebp-1c]
'0071ff0e    52                      push edx

' *** Reference to "__vbaI4Str"
'0071ff0f    ff1550124000            call dword ptr [00401250]
'0071ff15    8d8d2cffffff            lea ecx, dword ptr [ebp+ffffff2c]
'0071ff1b    51                      push ecx
'0071ff1c    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'0071ff22    52                      push edx
'0071ff23    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'0071ff29    8b06                    mov eax, dword ptr [esi]
'0071ff2b    56                      push esi
'0071ff2c    ff9000070000            call dword ptr [eax+00000700]
'0071ff32    85c0                    test ax, ax
'0071ff34    7d12                    jge 71ff48
'0071ff36    6800070000              push 00000700
'0071ff3b    68b0154300              push 004315b0
'0071ff40    56                      push esi
'0071ff41    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071ff42    ff1580104000            call dword ptr [00401080]
'0071ff48    8b4de8                  mov ecx, dword ptr [ebp-18]
'0071ff4b    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'0071ff51    8b18                    mov ebx, dword ptr [eax]
'0071ff53    51                      push ecx
'0071ff54    6848724500              push 00457248

' *** Reference to "__vbaStrCat"
'0071ff59    ff1570104000            call dword ptr [00401070]
var_5 = (vbNullString) & ("Liquidite disponible : ")
'0071ff5f    8bd0                    mov edx, eax
'0071ff61    8d4de0                  lea ecx, dword ptr [ebp-20]
'0071ff64    ffd7                    call edi
'0071ff66    8b952cffffff            mov edx, dword ptr [ebp+ffffff2c]
'0071ff6c    50                      push eax
'0071ff6d    52                      push edx

' *** Reference to "__vbaStrI4"
'0071ff6e    ff1520104000            call dword ptr [00401020]
'0071ff74    8bd0                    mov edx, eax
'0071ff76    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0071ff79    ffd7                    call edi
'0071ff7b    50                      push eax

' *** Reference to "__vbaStrCat"
'0071ff7c    ff1570104000            call dword ptr [00401070]
var_79 = (var_5) & (CStr(0))
'0071ff82    8bd0                    mov edx, eax
'0071ff84    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0071ff87    ffd7                    call edi
'0071ff89    50                      push eax
'0071ff8a    683c724500              push 0045723c

' *** Reference to "__vbaStrCat"
'0071ff8f    ff1570104000            call dword ptr [00401070]
var_2071 = (var_79) & (" Po")
'0071ff95    8bd0                    mov edx, eax
'0071ff97    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0071ff9a    ffd7                    call edi
'0071ff9c    50                      push eax
'0071ff9d    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'0071ffa2    ff1570104000            call dword ptr [00401070]
var_66 = (var_2071) & (vbCrLf)
'0071ffa8    8bd0                    mov edx, eax
'0071ffaa    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0071ffad    ffd7                    call edi
'0071ffaf    50                      push eax
'0071ffb0    8bc3                    mov eax, ebx
'0071ffb2    8b9d14ffffff            mov ebx, dword ptr [ebp+ffffff14]
'0071ffb8    53                      push ebx
'0071ffb9    ff90a4000000            call dword ptr [eax+000000a4]
'0071ffbf    dbe2                    fnclex
'0071ffc1    85c0                    test ax, ax
'0071ffc3    7d12                    jge 71ffd7
'0071ffc5    68a4000000              push 000000a4
'0071ffca    68d00d4300              push 00430dd0
'0071ffcf    53                      push ebx
'0071ffd0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071ffd1    ff1580104000            call dword ptr [00401080]
'0071ffd7    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0071ffda    51                      push ecx
'0071ffdb    8d55d4                  lea edx, dword ptr [ebp-2c]
'0071ffde    52                      push edx
'0071ffdf    8d45d8                  lea eax, dword ptr [ebp-28]
'0071ffe2    50                      push eax
'0071ffe3    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0071ffe6    51                      push ecx
'0071ffe7    8d55e0                  lea edx, dword ptr [ebp-20]
'0071ffea    52                      push edx
'0071ffeb    8d45e4                  lea eax, dword ptr [ebp-1c]
'0071ffee    50                      push eax
'0071ffef    8d4de8                  lea ecx, dword ptr [ebp-18]
'0071fff2    51                      push ecx
'0071fff3    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0071fff5    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, 0, -4620, -4624, -4628, -4632, -4636)
'0071fffb    8d55b8                  lea edx, dword ptr [ebp-48]
'0071fffe    52                      push edx
'0071ffff    8d45bc                  lea eax, dword ptr [ebp-44]
'00720002    50                      push eax
'00720003    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00720006    51                      push ecx
'00720007    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00720009    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58, var_61)
'0072000f    8b16                    mov edx, dword ptr [esi]
'00720011    83c430                  add esp, 30
'00720014    56                      push esi
'00720015    ff9204030000            call dword ptr [edx+00000304]

' *** Reference to "__vbaObjSet"
'0072001b    8b1db4104000            mov ebx, dword ptr [004010b4]
'00720021    50                      push eax
'00720022    8d45bc                  lea eax, dword ptr [ebp-44]
'00720025    50                      push eax
'00720026    ffd3                    call ebx
Set var_58 = 
'00720028    8b0e                    mov ecx, dword ptr [esi]
'0072002a    56                      push esi
'0072002b    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'00720031    ff9104030000            call dword ptr [ecx+00000304]
'00720037    50                      push eax
'00720038    8d55c0                  lea edx, dword ptr [ebp-40]
'0072003b    52                      push edx
'0072003c    ffd3                    call ebx
Set var_5 = var_58
'0072003e    8d4de8                  lea ecx, dword ptr [ebp-18]
'00720041    8bd8                    mov ebx, eax
'00720043    8b03                    mov eax, dword ptr [ebx]
'00720045    51                      push ecx
'00720046    53                      push ebx
'00720047    ff90a0000000            call dword ptr [eax+000000a0]
'0072004d    dbe2                    fnclex
'0072004f    85c0                    test ax, ax
'00720051    7d12                    jge 720065
'00720053    68a0000000              push 000000a0
'00720058    68d00d4300              push 00430dd0
'0072005d    53                      push ebx
'0072005e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072005f    ff1580104000            call dword ptr [00401080]
'00720065    8b45e8                  mov eax, dword ptr [ebp-18]
'00720068    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'0072006e    8b1a                    mov ebx, dword ptr [edx]
'00720070    50                      push eax
'00720071    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'00720076    ff1570104000            call dword ptr [00401070]
var_49 = (vbNullString) & (vbCrLf)
'0072007c    8bd0                    mov edx, eax
'0072007e    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00720081    ffd7                    call edi
'00720083    50                      push eax
'00720084    687c724500              push 0045727c

' *** Reference to "__vbaStrCat"
'00720089    ff1570104000            call dword ptr [00401070]
var_68 = (var_49) & ("Les revenus")
'0072008f    8bd0                    mov edx, eax
'00720091    8d4de0                  lea ecx, dword ptr [ebp-20]
'00720094    ffd7                    call edi
'00720096    50                      push eax
'00720097    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'0072009c    ff1570104000            call dword ptr [00401070]
var_2062 = (var_68) & (vbCrLf)
'007200a2    8bd0                    mov edx, eax
'007200a4    8d4ddc                  lea ecx, dword ptr [ebp-24]
'007200a7    ffd7                    call edi
'007200a9    8bcb                    mov ecx, ebx
'007200ab    8b9d20ffffff            mov ebx, dword ptr [ebp+ffffff20]
'007200b1    50                      push eax
'007200b2    53                      push ebx
'007200b3    ff91a4000000            call dword ptr [ecx+000000a4]
'007200b9    dbe2                    fnclex
'007200bb    85c0                    test ax, ax
'007200bd    7d12                    jge 7200d1
'007200bf    68a4000000              push 000000a4
'007200c4    68d00d4300              push 00430dd0
'007200c9    53                      push ebx
'007200ca    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007200cb    ff1580104000            call dword ptr [00401080]
'007200d1    8d55dc                  lea edx, dword ptr [ebp-24]
'007200d4    52                      push edx
'007200d5    8d45e0                  lea eax, dword ptr [ebp-20]
'007200d8    50                      push eax
'007200d9    8d4de4                  lea ecx, dword ptr [ebp-1c]
'007200dc    51                      push ecx
'007200dd    8d55e8                  lea edx, dword ptr [ebp-18]
'007200e0    52                      push edx
'007200e1    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'007200e3    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4640, -4644, -4648)
'007200e9    8d45bc                  lea eax, dword ptr [ebp-44]
'007200ec    50                      push eax
'007200ed    8d4dc0                  lea ecx, dword ptr [ebp-40]
'007200f0    51                      push ecx
'007200f1    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'007200f3    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58)
'007200f9    8b16                    mov edx, dword ptr [esi]
'007200fb    83c420                  add esp, 20
'007200fe    56                      push esi
'007200ff    ff9204030000            call dword ptr [edx+00000304]

' *** Reference to "__vbaObjSet"
'00720105    8b1db4104000            mov ebx, dword ptr [004010b4]
'0072010b    50                      push eax
'0072010c    8d45b8                  lea eax, dword ptr [ebp-48]
'0072010f    50                      push eax
'00720110    ffd3                    call ebx
Set var_61 = 
'00720112    8b0e                    mov ecx, dword ptr [esi]
'00720114    56                      push esi
'00720115    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'0072011b    ff9104030000            call dword ptr [ecx+00000304]
'00720121    50                      push eax
'00720122    8d55c0                  lea edx, dword ptr [ebp-40]
'00720125    52                      push edx
'00720126    ffd3                    call ebx
Set var_5 = var_61
'00720128    8d4de8                  lea ecx, dword ptr [ebp-18]
'0072012b    8bd8                    mov ebx, eax
'0072012d    8b03                    mov eax, dword ptr [ebx]
'0072012f    51                      push ecx
'00720130    53                      push ebx
'00720131    ff90a0000000            call dword ptr [eax+000000a0]
'00720137    dbe2                    fnclex
'00720139    85c0                    test ax, ax
'0072013b    7d12                    jge 72014f
'0072013d    68a0000000              push 000000a0
'00720142    68d00d4300              push 00430dd0
'00720147    53                      push ebx
'00720148    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720149    ff1580104000            call dword ptr [00401080]
'0072014f    8b16                    mov edx, dword ptr [esi]
'00720151    56                      push esi
'00720152    ff9218030000            call dword ptr [edx+00000318]
'00720158    50                      push eax
'00720159    8d45bc                  lea eax, dword ptr [ebp-44]
'0072015c    50                      push eax

' *** Reference to "__vbaObjSet"
'0072015d    ff15b4104000            call dword ptr [004010b4]
Set var_58 = var_58
'00720163    8d55e4                  lea edx, dword ptr [ebp-1c]
'00720166    8bd8                    mov ebx, eax
'00720168    8b0b                    mov ecx, dword ptr [ebx]
'0072016a    52                      push edx
'0072016b    53                      push ebx
'0072016c    ff91a0000000            call dword ptr [ecx+000000a0]
'00720172    dbe2                    fnclex
'00720174    85c0                    test ax, ax
'00720176    7d12                    jge 72018a
'00720178    68a0000000              push 000000a0
'0072017d    68d00d4300              push 00430dd0
'00720182    53                      push ebx
'00720183    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720184    ff1580104000            call dword ptr [00401080]
'0072018a    8b45e4                  mov eax, dword ptr [ebp-1c]
'0072018d    50                      push eax

' *** Reference to "__vbaI4Str"
'0072018e    ff1550124000            call dword ptr [00401250]
'00720194    8b0e                    mov ecx, dword ptr [esi]
'00720196    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'0072019c    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'007201a2    52                      push edx
'007201a3    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'007201a9    50                      push eax
'007201aa    56                      push esi
'007201ab    ff9108070000            call dword ptr [ecx+00000708]
'007201b1    85c0                    test ax, ax
'007201b3    7d12                    jge 7201c7
'007201b5    6808070000              push 00000708
'007201ba    68b0154300              push 004315b0
'007201bf    56                      push esi
'007201c0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007201c1    ff1580104000            call dword ptr [00401080]
'007201c7    8b55e8                  mov edx, dword ptr [ebp-18]
'007201ca    8b8d14ffffff            mov ecx, dword ptr [ebp+ffffff14]
'007201d0    8b19                    mov ebx, dword ptr [ecx]
'007201d2    52                      push edx
'007201d3    688c704500              push 0045708c

' *** Reference to "__vbaStrCat"
'007201d8    ff1570104000            call dword ptr [00401070]
var_23 = (vbNullString) & ("Revenus en or : ")
'007201de    8bd0                    mov edx, eax
'007201e0    8d4de0                  lea ecx, dword ptr [ebp-20]
'007201e3    ffd7                    call edi
'007201e5    50                      push eax
'007201e6    8b852cffffff            mov eax, dword ptr [ebp+ffffff2c]
'007201ec    50                      push eax

' *** Reference to "__vbaStrI4"
'007201ed    ff1520104000            call dword ptr [00401020]
'007201f3    8bd0                    mov edx, eax
'007201f5    8d4ddc                  lea ecx, dword ptr [ebp-24]
'007201f8    ffd7                    call edi
'007201fa    50                      push eax

' *** Reference to "__vbaStrCat"
'007201fb    ff1570104000            call dword ptr [00401070]
var_71 = (var_23) & (CStr(0))
'00720201    8bd0                    mov edx, eax
'00720203    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00720206    ffd7                    call edi
'00720208    50                      push eax
'00720209    68b4704500              push 004570b4

' *** Reference to "__vbaStrCat"
'0072020e    ff1570104000            call dword ptr [00401070]
var_873 = (var_71) & (" Po/mois")
'00720214    8bd0                    mov edx, eax
'00720216    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00720219    ffd7                    call edi
'0072021b    50                      push eax
'0072021c    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'00720221    ff1570104000            call dword ptr [00401070]
var_2583 = (var_873) & (vbCrLf)
'00720227    8bd0                    mov edx, eax
'00720229    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0072022c    ffd7                    call edi
'0072022e    8bcb                    mov ecx, ebx
'00720230    8b9d14ffffff            mov ebx, dword ptr [ebp+ffffff14]
'00720236    50                      push eax
'00720237    53                      push ebx
'00720238    ff91a4000000            call dword ptr [ecx+000000a4]
'0072023e    dbe2                    fnclex
'00720240    85c0                    test ax, ax
'00720242    7d12                    jge 720256
'00720244    68a4000000              push 000000a4
'00720249    68d00d4300              push 00430dd0
'0072024e    53                      push ebx
'0072024f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720250    ff1580104000            call dword ptr [00401080]
'00720256    8d55d0                  lea edx, dword ptr [ebp-30]
'00720259    52                      push edx
'0072025a    8d45d4                  lea eax, dword ptr [ebp-2c]
'0072025d    50                      push eax
'0072025e    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00720261    51                      push ecx
'00720262    8d55dc                  lea edx, dword ptr [ebp-24]
'00720265    52                      push edx
'00720266    8d45e0                  lea eax, dword ptr [ebp-20]
'00720269    50                      push eax
'0072026a    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0072026d    51                      push ecx
'0072026e    8d55e8                  lea edx, dword ptr [ebp-18]
'00720271    52                      push edx
'00720272    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'00720274    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4640, -4656, -4660, -4664, -4668, -4672)
'0072027a    8d45b8                  lea eax, dword ptr [ebp-48]
'0072027d    50                      push eax
'0072027e    8d4dbc                  lea ecx, dword ptr [ebp-44]
'00720281    51                      push ecx
'00720282    8d55c0                  lea edx, dword ptr [ebp-40]
'00720285    52                      push edx
'00720286    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'00720288    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58, var_61)
'0072028e    8b06                    mov eax, dword ptr [esi]
'00720290    83c430                  add esp, 30
'00720293    56                      push esi
'00720294    ff9004030000            call dword ptr [eax+00000304]

' *** Reference to "__vbaObjSet"
'0072029a    8b1db4104000            mov ebx, dword ptr [004010b4]
'007202a0    50                      push eax
'007202a1    8d4db8                  lea ecx, dword ptr [ebp-48]
'007202a4    51                      push ecx
'007202a5    ffd3                    call ebx
Set var_61 = Nothing
'007202a7    8b16                    mov edx, dword ptr [esi]
'007202a9    56                      push esi
'007202aa    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'007202b0    ff9204030000            call dword ptr [edx+00000304]
'007202b6    50                      push eax
'007202b7    8d45c0                  lea eax, dword ptr [ebp-40]
'007202ba    50                      push eax
'007202bb    ffd3                    call ebx
Set var_5 = Nothing
'007202bd    8d55e8                  lea edx, dword ptr [ebp-18]
'007202c0    8bd8                    mov ebx, eax
'007202c2    8b0b                    mov ecx, dword ptr [ebx]
'007202c4    52                      push edx
'007202c5    53                      push ebx
'007202c6    ff91a0000000            call dword ptr [ecx+000000a0]
'007202cc    dbe2                    fnclex
'007202ce    85c0                    test ax, ax
'007202d0    7d12                    jge 7202e4
'007202d2    68a0000000              push 000000a0
'007202d7    68d00d4300              push 00430dd0
'007202dc    53                      push ebx
'007202dd    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007202de    ff1580104000            call dword ptr [00401080]
'007202e4    8b06                    mov eax, dword ptr [esi]
'007202e6    56                      push esi
'007202e7    ff9018030000            call dword ptr [eax+00000318]
'007202ed    50                      push eax
'007202ee    8d4dbc                  lea ecx, dword ptr [ebp-44]
'007202f1    51                      push ecx

' *** Reference to "__vbaObjSet"
'007202f2    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Nothing
'007202f8    8bd8                    mov ebx, eax
'007202fa    8b13                    mov edx, dword ptr [ebx]
'007202fc    8d45e4                  lea eax, dword ptr [ebp-1c]
'007202ff    50                      push eax
'00720300    53                      push ebx
'00720301    ff92a0000000            call dword ptr [edx+000000a0]
'00720307    dbe2                    fnclex
'00720309    85c0                    test ax, ax
'0072030b    7d12                    jge 72031f
'0072030d    68a0000000              push 000000a0
'00720312    68d00d4300              push 00430dd0
'00720317    53                      push ebx
'00720318    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720319    ff1580104000            call dword ptr [00401080]
'0072031f    8b4de4                  mov ecx, dword ptr [ebp-1c]
'00720322    51                      push ecx

' *** Reference to "__vbaI4Str"
'00720323    ff1550124000            call dword ptr [00401250]
'00720329    8b16                    mov edx, dword ptr [esi]
'0072032b    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'00720331    8d852cffffff            lea eax, dword ptr [ebp+ffffff2c]
'00720337    50                      push eax
'00720338    8d8d30ffffff            lea ecx, dword ptr [ebp+ffffff30]
'0072033e    51                      push ecx
'0072033f    56                      push esi
'00720340    ff920c070000            call dword ptr [edx+0000070c]
'00720346    85c0                    test ax, ax
'00720348    7d12                    jge 72035c
'0072034a    680c070000              push 0000070c
'0072034f    68b0154300              push 004315b0
'00720354    56                      push esi
'00720355    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720356    ff1580104000            call dword ptr [00401080]
'0072035c    8b45e8                  mov eax, dword ptr [ebp-18]
'0072035f    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'00720365    8b1a                    mov ebx, dword ptr [edx]
'00720367    50                      push eax
'00720368    6898724500              push 00457298

' *** Reference to "__vbaStrCat"
'0072036d    ff1570104000            call dword ptr [00401070]
var_49 = (vbNullString) & ("Revenus en matières premières : ")
'00720373    8bd0                    mov edx, eax
'00720375    8d4de0                  lea ecx, dword ptr [ebp-20]
'00720378    ffd7                    call edi
'0072037a    8b8d2cffffff            mov ecx, dword ptr [ebp+ffffff2c]
'00720380    50                      push eax
'00720381    51                      push ecx

' *** Reference to "__vbaStrI4"
'00720382    ff1520104000            call dword ptr [00401020]
'00720388    8bd0                    mov edx, eax
'0072038a    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0072038d    ffd7                    call edi
'0072038f    50                      push eax

' *** Reference to "__vbaStrCat"
'00720390    ff1570104000            call dword ptr [00401070]
var_878 = (var_49) & (CStr(0))
'00720396    8bd0                    mov edx, eax
'00720398    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0072039b    ffd7                    call edi
'0072039d    50                      push eax
'0072039e    68b4704500              push 004570b4

' *** Reference to "__vbaStrCat"
'007203a3    ff1570104000            call dword ptr [00401070]
var_2475 = (var_878) & (" Po/mois")
'007203a9    8bd0                    mov edx, eax
'007203ab    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'007203ae    ffd7                    call edi
'007203b0    50                      push eax
'007203b1    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'007203b6    ff1570104000            call dword ptr [00401070]
var_879 = (var_2475) & (vbCrLf)
'007203bc    8bd0                    mov edx, eax
'007203be    8d4dd0                  lea ecx, dword ptr [ebp-30]
'007203c1    ffd7                    call edi
'007203c3    8bd3                    mov edx, ebx
'007203c5    8b9d14ffffff            mov ebx, dword ptr [ebp+ffffff14]
'007203cb    50                      push eax
'007203cc    53                      push ebx
'007203cd    ff92a4000000            call dword ptr [edx+000000a4]
'007203d3    dbe2                    fnclex
'007203d5    85c0                    test ax, ax
'007203d7    7d12                    jge 7203eb
'007203d9    68a4000000              push 000000a4
'007203de    68d00d4300              push 00430dd0
'007203e3    53                      push ebx
'007203e4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007203e5    ff1580104000            call dword ptr [00401080]
'007203eb    8d45d0                  lea eax, dword ptr [ebp-30]
'007203ee    50                      push eax
'007203ef    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'007203f2    51                      push ecx
'007203f3    8d55d8                  lea edx, dword ptr [ebp-28]
'007203f6    52                      push edx
'007203f7    8d45dc                  lea eax, dword ptr [ebp-24]
'007203fa    50                      push eax
'007203fb    8d4de0                  lea ecx, dword ptr [ebp-20]
'007203fe    51                      push ecx
'007203ff    8d55e4                  lea edx, dword ptr [ebp-1c]
'00720402    52                      push edx
'00720403    8d45e8                  lea eax, dword ptr [ebp-18]
'00720406    50                      push eax
'00720407    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'00720409    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4640, -4680, -4684, -4688, -4692, -4696)
'0072040f    8d4db8                  lea ecx, dword ptr [ebp-48]
'00720412    51                      push ecx
'00720413    8d55bc                  lea edx, dword ptr [ebp-44]
'00720416    52                      push edx
'00720417    8d45c0                  lea eax, dword ptr [ebp-40]
'0072041a    50                      push eax
'0072041b    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0072041d    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58, var_61)
'00720423    8b0e                    mov ecx, dword ptr [esi]
'00720425    83c430                  add esp, 30
'00720428    56                      push esi
'00720429    ff9104030000            call dword ptr [ecx+00000304]

' *** Reference to "__vbaObjSet"
'0072042f    8b1db4104000            mov ebx, dword ptr [004010b4]
'00720435    50                      push eax
'00720436    8d55b8                  lea edx, dword ptr [ebp-48]
'00720439    52                      push edx
'0072043a    ffd3                    call ebx
Set var_61 = 
'0072043c    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'00720442    8b06                    mov eax, dword ptr [esi]
'00720444    56                      push esi
'00720445    ff9004030000            call dword ptr [eax+00000304]
'0072044b    50                      push eax
'0072044c    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0072044f    51                      push ecx
'00720450    ffd3                    call ebx
Set var_5 = Nothing
'00720452    8bd8                    mov ebx, eax
'00720454    8b13                    mov edx, dword ptr [ebx]
'00720456    8d45e8                  lea eax, dword ptr [ebp-18]
'00720459    50                      push eax
'0072045a    53                      push ebx
'0072045b    ff92a0000000            call dword ptr [edx+000000a0]
'00720461    dbe2                    fnclex
'00720463    85c0                    test ax, ax
'00720465    7d12                    jge 720479
'00720467    68a0000000              push 000000a0
'0072046c    68d00d4300              push 00430dd0
'00720471    53                      push ebx
'00720472    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720473    ff1580104000            call dword ptr [00401080]
'00720479    8b0e                    mov ecx, dword ptr [esi]
'0072047b    56                      push esi
'0072047c    ff9118030000            call dword ptr [ecx+00000318]
'00720482    50                      push eax
'00720483    8d55bc                  lea edx, dword ptr [ebp-44]
'00720486    52                      push edx

' *** Reference to "__vbaObjSet"
'00720487    ff15b4104000            call dword ptr [004010b4]
Set var_58 = 
'0072048d    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00720490    8bd8                    mov ebx, eax
'00720492    8b03                    mov eax, dword ptr [ebx]
'00720494    51                      push ecx
'00720495    53                      push ebx
'00720496    ff90a0000000            call dword ptr [eax+000000a0]
'0072049c    dbe2                    fnclex
'0072049e    85c0                    test ax, ax
'007204a0    7d12                    jge 7204b4
'007204a2    68a0000000              push 000000a0
'007204a7    68d00d4300              push 00430dd0
'007204ac    53                      push ebx
'007204ad    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007204ae    ff1580104000            call dword ptr [00401080]
'007204b4    8b55e4                  mov edx, dword ptr [ebp-1c]
'007204b7    52                      push edx

' *** Reference to "__vbaI4Str"
'007204b8    ff1550124000            call dword ptr [00401250]
'007204be    8d8d2cffffff            lea ecx, dword ptr [ebp+ffffff2c]
'007204c4    51                      push ecx
'007204c5    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'007204cb    52                      push edx
'007204cc    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'007204d2    8b06                    mov eax, dword ptr [esi]
'007204d4    56                      push esi
'007204d5    ff9004070000            call dword ptr [eax+00000704]
'007204db    85c0                    test ax, ax
'007204dd    7d12                    jge 7204f1
'007204df    6804070000              push 00000704
'007204e4    68b0154300              push 004315b0
'007204e9    56                      push esi
'007204ea    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007204eb    ff1580104000            call dword ptr [00401080]
'007204f1    8b4de8                  mov ecx, dword ptr [ebp-18]
'007204f4    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'007204fa    8b18                    mov ebx, dword ptr [eax]
'007204fc    51                      push ecx
'007204fd    68e0724500              push 004572e0

' *** Reference to "__vbaStrCat"
'00720502    ff1570104000            call dword ptr [00401070]
var_2589 = (vbNullString) & ("Revenus totaux : ")
'00720508    8bd0                    mov edx, eax
'0072050a    8d4de0                  lea ecx, dword ptr [ebp-20]
'0072050d    ffd7                    call edi
'0072050f    8b952cffffff            mov edx, dword ptr [ebp+ffffff2c]
'00720515    50                      push eax
'00720516    52                      push edx

' *** Reference to "__vbaStrI4"
'00720517    ff1520104000            call dword ptr [00401020]
'0072051d    8bd0                    mov edx, eax
'0072051f    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00720522    ffd7                    call edi
'00720524    50                      push eax

' *** Reference to "__vbaStrCat"
'00720525    ff1570104000            call dword ptr [00401070]
var_882 = (var_2589) & (CStr(0))
'0072052b    8bd0                    mov edx, eax
'0072052d    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00720530    ffd7                    call edi
'00720532    50                      push eax
'00720533    68b4704500              push 004570b4

' *** Reference to "__vbaStrCat"
'00720538    ff1570104000            call dword ptr [00401070]
var_296 = (var_882) & (" Po/mois")
'0072053e    8bd0                    mov edx, eax
'00720540    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00720543    ffd7                    call edi
'00720545    50                      push eax
'00720546    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'0072054b    ff1570104000            call dword ptr [00401070]
var_883 = (var_296) & (vbCrLf)
'00720551    8bd0                    mov edx, eax
'00720553    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00720556    ffd7                    call edi
'00720558    50                      push eax
'00720559    8bc3                    mov eax, ebx
'0072055b    8b9d14ffffff            mov ebx, dword ptr [ebp+ffffff14]
'00720561    53                      push ebx
'00720562    ff90a4000000            call dword ptr [eax+000000a4]
'00720568    dbe2                    fnclex
'0072056a    85c0                    test ax, ax
'0072056c    7d12                    jge 720580
'0072056e    68a4000000              push 000000a4
'00720573    68d00d4300              push 00430dd0
'00720578    53                      push ebx
'00720579    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072057a    ff1580104000            call dword ptr [00401080]
'00720580    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00720583    51                      push ecx
'00720584    8d55d4                  lea edx, dword ptr [ebp-2c]
'00720587    52                      push edx
'00720588    8d45d8                  lea eax, dword ptr [ebp-28]
'0072058b    50                      push eax
'0072058c    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0072058f    51                      push ecx
'00720590    8d55e0                  lea edx, dword ptr [ebp-20]
'00720593    52                      push edx
'00720594    8d45e4                  lea eax, dword ptr [ebp-1c]
'00720597    50                      push eax
'00720598    8d4de8                  lea ecx, dword ptr [ebp-18]
'0072059b    51                      push ecx
'0072059c    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0072059e    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4640, -4704, -4708, -4712, -4716, -4720)
'007205a4    8d55b8                  lea edx, dword ptr [ebp-48]
'007205a7    52                      push edx
'007205a8    8d45bc                  lea eax, dword ptr [ebp-44]
'007205ab    50                      push eax
'007205ac    8d4dc0                  lea ecx, dword ptr [ebp-40]
'007205af    51                      push ecx
'007205b0    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'007205b2    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58, var_61)
'007205b8    8b16                    mov edx, dword ptr [esi]
'007205ba    83c430                  add esp, 30
'007205bd    56                      push esi
'007205be    ff9204030000            call dword ptr [edx+00000304]

' *** Reference to "__vbaObjSet"
'007205c4    8b1db4104000            mov ebx, dword ptr [004010b4]
'007205ca    50                      push eax
'007205cb    8d45bc                  lea eax, dword ptr [ebp-44]
'007205ce    50                      push eax
'007205cf    ffd3                    call ebx
Set var_58 = 
'007205d1    8b0e                    mov ecx, dword ptr [esi]
'007205d3    56                      push esi
'007205d4    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'007205da    ff9104030000            call dword ptr [ecx+00000304]
'007205e0    50                      push eax
'007205e1    8d55c0                  lea edx, dword ptr [ebp-40]
'007205e4    52                      push edx
'007205e5    ffd3                    call ebx
Set var_5 = var_58
'007205e7    8d4de8                  lea ecx, dword ptr [ebp-18]
'007205ea    8bd8                    mov ebx, eax
'007205ec    8b03                    mov eax, dword ptr [ebx]
'007205ee    51                      push ecx
'007205ef    53                      push ebx
'007205f0    ff90a0000000            call dword ptr [eax+000000a0]
'007205f6    dbe2                    fnclex
'007205f8    85c0                    test ax, ax
'007205fa    7d12                    jge 72060e
'007205fc    68a0000000              push 000000a0
'00720601    68d00d4300              push 00430dd0
'00720606    53                      push ebx
'00720607    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720608    ff1580104000            call dword ptr [00401080]
'0072060e    8b45e8                  mov eax, dword ptr [ebp-18]
'00720611    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'00720617    8b1a                    mov ebx, dword ptr [edx]
'00720619    50                      push eax
'0072061a    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'0072061f    ff1570104000            call dword ptr [00401070]
var_49 = (vbNullString) & (vbCrLf)
'00720625    8bd0                    mov edx, eax
'00720627    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0072062a    ffd7                    call edi
'0072062c    50                      push eax
'0072062d    6808734500              push 00457308

' *** Reference to "__vbaStrCat"
'00720632    ff1570104000            call dword ptr [00401070]
var_2445 = (var_49) & ("Les impôts")
'00720638    8bd0                    mov edx, eax
'0072063a    8d4de0                  lea ecx, dword ptr [ebp-20]
'0072063d    ffd7                    call edi
'0072063f    50                      push eax
'00720640    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'00720645    ff1570104000            call dword ptr [00401070]
var_886 = (var_2445) & (vbCrLf)
'0072064b    8bd0                    mov edx, eax
'0072064d    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00720650    ffd7                    call edi
'00720652    8bcb                    mov ecx, ebx
'00720654    8b9d20ffffff            mov ebx, dword ptr [ebp+ffffff20]
'0072065a    50                      push eax
'0072065b    53                      push ebx
'0072065c    ff91a4000000            call dword ptr [ecx+000000a4]
'00720662    dbe2                    fnclex
'00720664    85c0                    test ax, ax
'00720666    7d12                    jge 72067a
'00720668    68a4000000              push 000000a4
'0072066d    68d00d4300              push 00430dd0
'00720672    53                      push ebx
'00720673    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720674    ff1580104000            call dword ptr [00401080]
'0072067a    8d55dc                  lea edx, dword ptr [ebp-24]
'0072067d    52                      push edx
'0072067e    8d45e0                  lea eax, dword ptr [ebp-20]
'00720681    50                      push eax
'00720682    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00720685    51                      push ecx
'00720686    8d55e8                  lea edx, dword ptr [ebp-18]
'00720689    52                      push edx
'0072068a    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'0072068c    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4724, -4728, -4732)
'00720692    8d45bc                  lea eax, dword ptr [ebp-44]
'00720695    50                      push eax
'00720696    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00720699    51                      push ecx
'0072069a    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0072069c    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58)
'007206a2    8b16                    mov edx, dword ptr [esi]
'007206a4    83c420                  add esp, 20
'007206a7    56                      push esi
'007206a8    ff9218030000            call dword ptr [edx+00000318]
'007206ae    50                      push eax
'007206af    8d45bc                  lea eax, dword ptr [ebp-44]
'007206b2    50                      push eax

' *** Reference to "__vbaObjSet"
'007206b3    ff15b4104000            call dword ptr [004010b4]
Set var_58 = 
'007206b9    8d55e4                  lea edx, dword ptr [ebp-1c]
'007206bc    8bd8                    mov ebx, eax
'007206be    8b0b                    mov ecx, dword ptr [ebx]
'007206c0    52                      push edx
'007206c1    53                      push ebx
'007206c2    ff91a0000000            call dword ptr [ecx+000000a0]
'007206c8    dbe2                    fnclex
'007206ca    85c0                    test ax, ax
'007206cc    7d12                    jge 7206e0
'007206ce    68a0000000              push 000000a0
'007206d3    68d00d4300              push 00430dd0
'007206d8    53                      push ebx
'007206d9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007206da    ff1580104000            call dword ptr [00401080]
'007206e0    8b06                    mov eax, dword ptr [esi]
'007206e2    56                      push esi
'007206e3    ff900c030000            call dword ptr [eax+0000030c]
'007206e9    50                      push eax
'007206ea    8d4db8                  lea ecx, dword ptr [ebp-48]
'007206ed    51                      push ecx

' *** Reference to "__vbaObjSet"
'007206ee    ff15b4104000            call dword ptr [004010b4]
Set var_61 = Nothing
'007206f4    8bd8                    mov ebx, eax
'007206f6    8b13                    mov edx, dword ptr [ebx]
'007206f8    8d45e0                  lea eax, dword ptr [ebp-20]
'007206fb    50                      push eax
'007206fc    53                      push ebx
'007206fd    ff92a0000000            call dword ptr [edx+000000a0]
'00720703    dbe2                    fnclex
'00720705    85c0                    test ax, ax
'00720707    7d12                    jge 72071b
'00720709    68a0000000              push 000000a0
'0072070e    68d00d4300              push 00430dd0
'00720713    53                      push ebx
'00720714    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720715    ff1580104000            call dword ptr [00401080]
'0072071b    8b0e                    mov ecx, dword ptr [esi]
'0072071d    56                      push esi
'0072071e    ff9104030000            call dword ptr [ecx+00000304]

' *** Reference to "__vbaObjSet"
'00720724    8b1db4104000            mov ebx, dword ptr [004010b4]
'0072072a    50                      push eax
'0072072b    8d55b4                  lea edx, dword ptr [ebp-4c]
'0072072e    52                      push edx
'0072072f    ffd3                    call ebx
Set var_62 = 
'00720731    89850cffffff            mov dword ptr [ebp+ffffff0c], eax
'00720737    8b06                    mov eax, dword ptr [esi]
'00720739    56                      push esi
'0072073a    ff9004030000            call dword ptr [eax+00000304]
'00720740    50                      push eax
'00720741    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00720744    51                      push ecx
'00720745    ffd3                    call ebx
Set var_5 = Nothing
'00720747    8bd8                    mov ebx, eax
'00720749    8b13                    mov edx, dword ptr [ebx]
'0072074b    8d45e8                  lea eax, dword ptr [ebp-18]
'0072074e    50                      push eax
'0072074f    53                      push ebx
'00720750    ff92a0000000            call dword ptr [edx+000000a0]
'00720756    dbe2                    fnclex
'00720758    85c0                    test ax, ax
'0072075a    7d12                    jge 72076e
'0072075c    68a0000000              push 000000a0
'00720761    68d00d4300              push 00430dd0
'00720766    53                      push ebx
'00720767    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720768    ff1580104000            call dword ptr [00401080]
'0072076e    8b4de4                  mov ecx, dword ptr [ebp-1c]
'00720771    51                      push ecx

' *** Reference to "__vbaI4Str"
'00720772    ff1550124000            call dword ptr [00401250]
'00720778    8b16                    mov edx, dword ptr [esi]
'0072077a    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'00720780    8d852cffffff            lea eax, dword ptr [ebp+ffffff2c]
'00720786    50                      push eax
'00720787    8d8d30ffffff            lea ecx, dword ptr [ebp+ffffff30]
'0072078d    51                      push ecx
'0072078e    56                      push esi
'0072078f    ff9208070000            call dword ptr [edx+00000708]
'00720795    85c0                    test ax, ax
'00720797    7d12                    jge 7207ab
'00720799    6808070000              push 00000708
'0072079e    68b0154300              push 004315b0
'007207a3    56                      push esi
'007207a4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007207a5    ff1580104000            call dword ptr [00401080]
'007207ab    8b45e8                  mov eax, dword ptr [ebp-18]
'007207ae    8b950cffffff            mov edx, dword ptr [ebp+ffffff0c]
'007207b4    8b1a                    mov ebx, dword ptr [edx]
'007207b6    50                      push eax
'007207b7    6824734500              push 00457324

' *** Reference to "__vbaStrCat"
'007207bc    ff1570104000            call dword ptr [00401070]
var_49 = (vbNullString) & ("Impôts en or : ")
'007207c2    8bd0                    mov edx, eax
'007207c4    8d4ddc                  lea ecx, dword ptr [ebp-24]
'007207c7    ffd7                    call edi
'007207c9    8b4de0                  mov ecx, dword ptr [ebp-20]
'007207cc    50                      push eax
'007207cd    51                      push ecx

' *** Reference to "__vbaR8Str"
'007207ce    ff1524124000            call dword ptr [00401224]
'007207d4    db852cffffff            fild dword ptr [ebp+ffffff2c]
'007207da    83ec08                  sub esp, 08
'007207dd    dd9dd4feffff            fstp qword ptr [ebp+fffffed4]
'var_825 = (00)
'007207e3    dc8dd4feffff            fmul qword ptr [ebp+fffffed4]
'007207e9    833d00b0720000          cmp dword ptr [0072b000], 00
'007207f0    7508                    jne 7207fa

If (vbNullChar = 0) Then
'007207f2    dc35286c4000            fdiv qword ptr [00406c28]
'007207f8    eb11                    jmp 72080b
    
End If
'007207fa    ff352c6c4000            push dword ptr [00406c2c]
'00720800    ff35286c4000            push dword ptr [00406c28]

' *** Reference to "_adj_fdiv_m64"
'00720806    e8796aceff              call 407284
'0072080b    dfe0                    fnstsw ax
'0072080d    a80d                    test al, 0d
'0072080f    0f85e1130000            jne 721bf6
'00720815    dd1c24                  fstp qword ptr [esp]
'var_90 = (00)

' *** Reference to "__vbaStrR8"
'00720818    ff1580114000            call dword ptr [00401180]
'0072081e    8bd0                    mov edx, eax
'00720820    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00720823    ffd7                    call edi
'00720825    50                      push eax

' *** Reference to "__vbaStrCat"
'00720826    ff1570104000            call dword ptr [00401070]
var_2084 = (var_49) & (CStr(((CDbl(var_2445) * ()) / 100#)))
'0072082c    8bd0                    mov edx, eax
'0072082e    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00720831    ffd7                    call edi
'00720833    50                      push eax
'00720834    68b4704500              push 004570b4

' *** Reference to "__vbaStrCat"
'00720839    ff1570104000            call dword ptr [00401070]
var_2370 = (var_2084) & (" Po/mois")
'0072083f    8bd0                    mov edx, eax
'00720841    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00720844    ffd7                    call edi
'00720846    50                      push eax
'00720847    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'0072084c    ff1570104000            call dword ptr [00401070]
var_1503 = (var_2370) & (vbCrLf)
'00720852    8bd0                    mov edx, eax
'00720854    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00720857    ffd7                    call edi
'00720859    8bd3                    mov edx, ebx
'0072085b    8b9d0cffffff            mov ebx, dword ptr [ebp+ffffff0c]
'00720861    50                      push eax
'00720862    53                      push ebx
'00720863    ff92a4000000            call dword ptr [edx+000000a4]
'00720869    dbe2                    fnclex
'0072086b    85c0                    test ax, ax
'0072086d    7d12                    jge 720881
'0072086f    68a4000000              push 000000a4
'00720874    68d00d4300              push 00430dd0
'00720879    53                      push ebx
'0072087a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072087b    ff1580104000            call dword ptr [00401080]
'00720881    8d45cc                  lea eax, dword ptr [ebp-34]
'00720884    50                      push eax
'00720885    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00720888    51                      push ecx
'00720889    8d55d4                  lea edx, dword ptr [ebp-2c]
'0072088c    52                      push edx
'0072088d    8d45d8                  lea eax, dword ptr [ebp-28]
'00720890    50                      push eax
'00720891    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00720894    51                      push ecx
'00720895    8d55e0                  lea edx, dword ptr [ebp-20]
'00720898    52                      push edx
'00720899    8d45e4                  lea eax, dword ptr [ebp-1c]
'0072089c    50                      push eax
'0072089d    8d4de8                  lea ecx, dword ptr [ebp-18]
'007208a0    51                      push ecx
'007208a1    6a08                    push 08

' *** Reference to "__vbaFreeStrList"
'007208a3    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4724, -4728, -4740, -4744, -4748, -4752, -4756)
'007208a9    8d55b4                  lea edx, dword ptr [ebp-4c]
'007208ac    52                      push edx
'007208ad    8d45b8                  lea eax, dword ptr [ebp-48]
'007208b0    50                      push eax
'007208b1    8d4dbc                  lea ecx, dword ptr [ebp-44]
'007208b4    51                      push ecx
'007208b5    8d55c0                  lea edx, dword ptr [ebp-40]
'007208b8    52                      push edx
'007208b9    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'007208bb    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58, var_61, var_62)
'007208c1    8b06                    mov eax, dword ptr [esi]
'007208c3    83c438                  add esp, 38
'007208c6    56                      push esi
'007208c7    ff9018030000            call dword ptr [eax+00000318]
'007208cd    50                      push eax
'007208ce    8d4dbc                  lea ecx, dword ptr [ebp-44]
'007208d1    51                      push ecx

' *** Reference to "__vbaObjSet"
'007208d2    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Nothing
'007208d8    8bd8                    mov ebx, eax
'007208da    8b13                    mov edx, dword ptr [ebx]
'007208dc    8d45e4                  lea eax, dword ptr [ebp-1c]
'007208df    50                      push eax
'007208e0    53                      push ebx
'007208e1    ff92a0000000            call dword ptr [edx+000000a0]
'007208e7    dbe2                    fnclex
'007208e9    85c0                    test ax, ax
'007208eb    7d12                    jge 7208ff
'007208ed    68a0000000              push 000000a0
'007208f2    68d00d4300              push 00430dd0
'007208f7    53                      push ebx
'007208f8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007208f9    ff1580104000            call dword ptr [00401080]
'007208ff    8b0e                    mov ecx, dword ptr [esi]
'00720901    56                      push esi
'00720902    ff910c030000            call dword ptr [ecx+0000030c]
'00720908    50                      push eax
'00720909    8d55b8                  lea edx, dword ptr [ebp-48]
'0072090c    52                      push edx

' *** Reference to "__vbaObjSet"
'0072090d    ff15b4104000            call dword ptr [004010b4]
Set var_61 = 
'00720913    8d4de0                  lea ecx, dword ptr [ebp-20]
'00720916    8bd8                    mov ebx, eax
'00720918    8b03                    mov eax, dword ptr [ebx]
'0072091a    51                      push ecx
'0072091b    53                      push ebx
'0072091c    ff90a0000000            call dword ptr [eax+000000a0]
'00720922    dbe2                    fnclex
'00720924    85c0                    test ax, ax
'00720926    7d12                    jge 72093a
'00720928    68a0000000              push 000000a0
'0072092d    68d00d4300              push 00430dd0
'00720932    53                      push ebx
'00720933    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720934    ff1580104000            call dword ptr [00401080]
'0072093a    8b16                    mov edx, dword ptr [esi]
'0072093c    56                      push esi
'0072093d    ff9204030000            call dword ptr [edx+00000304]

' *** Reference to "__vbaObjSet"
'00720943    8b1db4104000            mov ebx, dword ptr [004010b4]
'00720949    50                      push eax
'0072094a    8d45b4                  lea eax, dword ptr [ebp-4c]
'0072094d    50                      push eax
'0072094e    ffd3                    call ebx
Set var_62 = -4724
'00720950    8b0e                    mov ecx, dword ptr [esi]
'00720952    56                      push esi
'00720953    89850cffffff            mov dword ptr [ebp+ffffff0c], eax
'00720959    ff9104030000            call dword ptr [ecx+00000304]
'0072095f    50                      push eax
'00720960    8d55c0                  lea edx, dword ptr [ebp-40]
'00720963    52                      push edx
'00720964    ffd3                    call ebx
Set var_5 = var_62
'00720966    8d4de8                  lea ecx, dword ptr [ebp-18]
'00720969    8bd8                    mov ebx, eax
'0072096b    8b03                    mov eax, dword ptr [ebx]
'0072096d    51                      push ecx
'0072096e    53                      push ebx
'0072096f    ff90a0000000            call dword ptr [eax+000000a0]
'00720975    dbe2                    fnclex
'00720977    85c0                    test ax, ax
'00720979    7d12                    jge 72098d
'0072097b    68a0000000              push 000000a0
'00720980    68d00d4300              push 00430dd0
'00720985    53                      push ebx
'00720986    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720987    ff1580104000            call dword ptr [00401080]
'0072098d    8b55e4                  mov edx, dword ptr [ebp-1c]
'00720990    52                      push edx

' *** Reference to "__vbaI4Str"
'00720991    ff1550124000            call dword ptr [00401250]
'00720997    8d8d2cffffff            lea ecx, dword ptr [ebp+ffffff2c]
'0072099d    51                      push ecx
'0072099e    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'007209a4    52                      push edx
'007209a5    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'007209ab    8b06                    mov eax, dword ptr [esi]
'007209ad    56                      push esi
'007209ae    ff900c070000            call dword ptr [eax+0000070c]
'007209b4    85c0                    test ax, ax
'007209b6    7d12                    jge 7209ca
'007209b8    680c070000              push 0000070c
'007209bd    68b0154300              push 004315b0
'007209c2    56                      push esi
'007209c3    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007209c4    ff1580104000            call dword ptr [00401080]
'007209ca    8b4de8                  mov ecx, dword ptr [ebp-18]
'007209cd    8b850cffffff            mov eax, dword ptr [ebp+ffffff0c]
'007209d3    8b18                    mov ebx, dword ptr [eax]
'007209d5    51                      push ecx
'007209d6    6848734500              push 00457348

' *** Reference to "__vbaStrCat"
'007209db    ff1570104000            call dword ptr [00401070]
var_2445 = (vbNullString) & ("Impôts en matières premières : ")
'007209e1    8bd0                    mov edx, eax
'007209e3    8d4ddc                  lea ecx, dword ptr [ebp-24]
'007209e6    ffd7                    call edi
'007209e8    8b55e0                  mov edx, dword ptr [ebp-20]
'007209eb    50                      push eax
'007209ec    52                      push edx

' *** Reference to "__vbaR8Str"
'007209ed    ff1524124000            call dword ptr [00401224]
'007209f3    db852cffffff            fild dword ptr [ebp+ffffff2c]
'007209f9    83ec08                  sub esp, 08
'007209fc    dd9dc8feffff            fstp qword ptr [ebp+fffffec8]
'var_65 = (00)
'00720a02    dc8dc8feffff            fmul qword ptr [ebp+fffffec8]
'00720a08    833d00b0720000          cmp dword ptr [0072b000], 00
'00720a0f    7508                    jne 720a19

If (vbNullChar = 0) Then
'00720a11    dc35286c4000            fdiv qword ptr [00406c28]
'00720a17    eb11                    jmp 720a2a
    
End If
'00720a19    ff352c6c4000            push dword ptr [00406c2c]
'00720a1f    ff35286c4000            push dword ptr [00406c28]

' *** Reference to "_adj_fdiv_m64"
'00720a25    e85a68ceff              call 407284
'00720a2a    dfe0                    fnstsw ax
'00720a2c    a80d                    test al, 0d
'00720a2e    0f85c2110000            jne 721bf6
'00720a34    dd1c24                  fstp qword ptr [esp]
'var_88 = (00)

' *** Reference to "__vbaStrR8"
'00720a37    ff1580114000            call dword ptr [00401180]
'00720a3d    8bd0                    mov edx, eax
'00720a3f    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00720a42    ffd7                    call edi
'00720a44    50                      push eax

' *** Reference to "__vbaStrCat"
'00720a45    ff1570104000            call dword ptr [00401070]
var_2189 = (var_2445) & (CStr(((CDbl(var_2445) * ()) / 100#)))
'00720a4b    8bd0                    mov edx, eax
'00720a4d    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00720a50    ffd7                    call edi
'00720a52    50                      push eax
'00720a53    68b4704500              push 004570b4

' *** Reference to "__vbaStrCat"
'00720a58    ff1570104000            call dword ptr [00401070]
var_1528 = (var_2189) & (" Po/mois")
'00720a5e    8bd0                    mov edx, eax
'00720a60    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00720a63    ffd7                    call edi
'00720a65    50                      push eax
'00720a66    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'00720a6b    ff1570104000            call dword ptr [00401070]
var_2124 = (var_1528) & (vbCrLf)
'00720a71    8bd0                    mov edx, eax
'00720a73    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00720a76    ffd7                    call edi
'00720a78    50                      push eax
'00720a79    8bc3                    mov eax, ebx
'00720a7b    8b9d0cffffff            mov ebx, dword ptr [ebp+ffffff0c]
'00720a81    53                      push ebx
'00720a82    ff90a4000000            call dword ptr [eax+000000a4]
'00720a88    dbe2                    fnclex
'00720a8a    85c0                    test ax, ax
'00720a8c    7d12                    jge 720aa0
'00720a8e    68a4000000              push 000000a4
'00720a93    68d00d4300              push 00430dd0
'00720a98    53                      push ebx
'00720a99    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720a9a    ff1580104000            call dword ptr [00401080]
'00720aa0    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00720aa3    51                      push ecx
'00720aa4    8d55d0                  lea edx, dword ptr [ebp-30]
'00720aa7    52                      push edx
'00720aa8    8d45d4                  lea eax, dword ptr [ebp-2c]
'00720aab    50                      push eax
'00720aac    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00720aaf    51                      push ecx
'00720ab0    8d55dc                  lea edx, dword ptr [ebp-24]
'00720ab3    52                      push edx
'00720ab4    8d45e0                  lea eax, dword ptr [ebp-20]
'00720ab7    50                      push eax
'00720ab8    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00720abb    51                      push ecx
'00720abc    8d55e8                  lea edx, dword ptr [ebp-18]
'00720abf    52                      push edx
'00720ac0    6a08                    push 08

' *** Reference to "__vbaFreeStrList"
'00720ac2    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4724, -4728, -4764, -4768, -4772, -4776, -4780)
'00720ac8    8d45b4                  lea eax, dword ptr [ebp-4c]
'00720acb    50                      push eax
'00720acc    8d4db8                  lea ecx, dword ptr [ebp-48]
'00720acf    51                      push ecx
'00720ad0    8d55bc                  lea edx, dword ptr [ebp-44]
'00720ad3    52                      push edx
'00720ad4    8d45c0                  lea eax, dword ptr [ebp-40]
'00720ad7    50                      push eax
'00720ad8    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'00720ada    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58, var_61, var_62)
'00720ae0    8b0e                    mov ecx, dword ptr [esi]
'00720ae2    83c438                  add esp, 38
'00720ae5    56                      push esi
'00720ae6    ff9118030000            call dword ptr [ecx+00000318]
'00720aec    50                      push eax
'00720aed    8d55bc                  lea edx, dword ptr [ebp-44]
'00720af0    52                      push edx

' *** Reference to "__vbaObjSet"
'00720af1    ff15b4104000            call dword ptr [004010b4]
Set var_58 = 
'00720af7    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00720afa    8bd8                    mov ebx, eax
'00720afc    8b03                    mov eax, dword ptr [ebx]
'00720afe    51                      push ecx
'00720aff    53                      push ebx
'00720b00    ff90a0000000            call dword ptr [eax+000000a0]
'00720b06    dbe2                    fnclex
'00720b08    85c0                    test ax, ax
'00720b0a    7d12                    jge 720b1e
'00720b0c    68a0000000              push 000000a0
'00720b11    68d00d4300              push 00430dd0
'00720b16    53                      push ebx
'00720b17    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720b18    ff1580104000            call dword ptr [00401080]
'00720b1e    8b16                    mov edx, dword ptr [esi]
'00720b20    56                      push esi
'00720b21    ff920c030000            call dword ptr [edx+0000030c]
'00720b27    50                      push eax
'00720b28    8d45b8                  lea eax, dword ptr [ebp-48]
'00720b2b    50                      push eax

' *** Reference to "__vbaObjSet"
'00720b2c    ff15b4104000            call dword ptr [004010b4]
Set var_61 = var_5
'00720b32    8d55e0                  lea edx, dword ptr [ebp-20]
'00720b35    8bd8                    mov ebx, eax
'00720b37    8b0b                    mov ecx, dword ptr [ebx]
'00720b39    52                      push edx
'00720b3a    53                      push ebx
'00720b3b    ff91a0000000            call dword ptr [ecx+000000a0]
'00720b41    dbe2                    fnclex
'00720b43    85c0                    test ax, ax
'00720b45    7d12                    jge 720b59
'00720b47    68a0000000              push 000000a0
'00720b4c    68d00d4300              push 00430dd0
'00720b51    53                      push ebx
'00720b52    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720b53    ff1580104000            call dword ptr [00401080]
'00720b59    8b06                    mov eax, dword ptr [esi]
'00720b5b    56                      push esi
'00720b5c    ff9004030000            call dword ptr [eax+00000304]

' *** Reference to "__vbaObjSet"
'00720b62    8b1db4104000            mov ebx, dword ptr [004010b4]
'00720b68    50                      push eax
'00720b69    8d4db4                  lea ecx, dword ptr [ebp-4c]
'00720b6c    51                      push ecx
'00720b6d    ffd3                    call ebx
Set var_62 = Nothing
'00720b6f    8b16                    mov edx, dword ptr [esi]
'00720b71    56                      push esi
'00720b72    89850cffffff            mov dword ptr [ebp+ffffff0c], eax
'00720b78    ff9204030000            call dword ptr [edx+00000304]
'00720b7e    50                      push eax
'00720b7f    8d45c0                  lea eax, dword ptr [ebp-40]
'00720b82    50                      push eax
'00720b83    ffd3                    call ebx
Set var_5 = Nothing
'00720b85    8d55e8                  lea edx, dword ptr [ebp-18]
'00720b88    8bd8                    mov ebx, eax
'00720b8a    8b0b                    mov ecx, dword ptr [ebx]
'00720b8c    52                      push edx
'00720b8d    53                      push ebx
'00720b8e    ff91a0000000            call dword ptr [ecx+000000a0]
'00720b94    dbe2                    fnclex
'00720b96    85c0                    test ax, ax
'00720b98    7d12                    jge 720bac
'00720b9a    68a0000000              push 000000a0
'00720b9f    68d00d4300              push 00430dd0
'00720ba4    53                      push ebx
'00720ba5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720ba6    ff1580104000            call dword ptr [00401080]
'00720bac    8b45e4                  mov eax, dword ptr [ebp-1c]
'00720baf    50                      push eax

' *** Reference to "__vbaI4Str"
'00720bb0    ff1550124000            call dword ptr [00401250]
'00720bb6    8b0e                    mov ecx, dword ptr [esi]
'00720bb8    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'00720bbe    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'00720bc4    52                      push edx
'00720bc5    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'00720bcb    50                      push eax
'00720bcc    56                      push esi
'00720bcd    ff9104070000            call dword ptr [ecx+00000704]
'00720bd3    85c0                    test ax, ax
'00720bd5    7d12                    jge 720be9
'00720bd7    6804070000              push 00000704
'00720bdc    68b0154300              push 004315b0
'00720be1    56                      push esi
'00720be2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720be3    ff1580104000            call dword ptr [00401080]
'00720be9    8b55e8                  mov edx, dword ptr [ebp-18]
'00720bec    8b8d0cffffff            mov ecx, dword ptr [ebp+ffffff0c]
'00720bf2    8b19                    mov ebx, dword ptr [ecx]
'00720bf4    52                      push edx
'00720bf5    688c734500              push 0045738c

' *** Reference to "__vbaStrCat"
'00720bfa    ff1570104000            call dword ptr [00401070]
var_23 = (vbNullString) & ("Impôts totaux : ")
'00720c00    8bd0                    mov edx, eax
'00720c02    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00720c05    ffd7                    call edi
'00720c07    50                      push eax
'00720c08    8b45e0                  mov eax, dword ptr [ebp-20]
'00720c0b    50                      push eax

' *** Reference to "__vbaR8Str"
'00720c0c    ff1524124000            call dword ptr [00401224]
'00720c12    db852cffffff            fild dword ptr [ebp+ffffff2c]
'00720c18    83ec08                  sub esp, 08
'00720c1b    dd9dbcfeffff            fstp qword ptr [ebp+fffffebc]
'var_482 = (00)
'00720c21    dc8dbcfeffff            fmul qword ptr [ebp+fffffebc]
'00720c27    833d00b0720000          cmp dword ptr [0072b000], 00
'00720c2e    7508                    jne 720c38

If (vbNullChar = 0) Then
'00720c30    dc35286c4000            fdiv qword ptr [00406c28]
'00720c36    eb11                    jmp 720c49
    
End If
'00720c38    ff352c6c4000            push dword ptr [00406c2c]
'00720c3e    ff35286c4000            push dword ptr [00406c28]

' *** Reference to "_adj_fdiv_m64"
'00720c44    e83b66ceff              call 407284
'00720c49    dfe0                    fnstsw ax
'00720c4b    a80d                    test al, 0d
'00720c4d    0f85a30f0000            jne 721bf6
'00720c53    dd1c24                  fstp qword ptr [esp]
'var_59 = (00)

' *** Reference to "__vbaStrR8"
'00720c56    ff1580114000            call dword ptr [00401180]
'00720c5c    8bd0                    mov edx, eax
'00720c5e    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00720c61    ffd7                    call edi
'00720c63    50                      push eax

' *** Reference to "__vbaStrCat"
'00720c64    ff1570104000            call dword ptr [00401070]
var_2437 = (var_23) & (CStr(((CDbl(var_2445) * ()) / 100#)))
'00720c6a    8bd0                    mov edx, eax
'00720c6c    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00720c6f    ffd7                    call edi
'00720c71    50                      push eax
'00720c72    68b4704500              push 004570b4

' *** Reference to "__vbaStrCat"
'00720c77    ff1570104000            call dword ptr [00401070]
var_2436 = (var_2437) & (" Po/mois")
'00720c7d    8bd0                    mov edx, eax
'00720c7f    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00720c82    ffd7                    call edi
'00720c84    50                      push eax
'00720c85    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'00720c8a    ff1570104000            call dword ptr [00401070]
var_505 = (var_2436) & (vbCrLf)
'00720c90    8bd0                    mov edx, eax
'00720c92    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00720c95    ffd7                    call edi
'00720c97    8bcb                    mov ecx, ebx
'00720c99    8b9d0cffffff            mov ebx, dword ptr [ebp+ffffff0c]
'00720c9f    50                      push eax
'00720ca0    53                      push ebx
'00720ca1    ff91a4000000            call dword ptr [ecx+000000a4]
'00720ca7    dbe2                    fnclex
'00720ca9    85c0                    test ax, ax
'00720cab    7d12                    jge 720cbf
'00720cad    68a4000000              push 000000a4
'00720cb2    68d00d4300              push 00430dd0
'00720cb7    53                      push ebx
'00720cb8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720cb9    ff1580104000            call dword ptr [00401080]
'00720cbf    8d55cc                  lea edx, dword ptr [ebp-34]
'00720cc2    52                      push edx
'00720cc3    8d45d0                  lea eax, dword ptr [ebp-30]
'00720cc6    50                      push eax
'00720cc7    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00720cca    51                      push ecx
'00720ccb    8d55d8                  lea edx, dword ptr [ebp-28]
'00720cce    52                      push edx
'00720ccf    8d45dc                  lea eax, dword ptr [ebp-24]
'00720cd2    50                      push eax
'00720cd3    8d4de0                  lea ecx, dword ptr [ebp-20]
'00720cd6    51                      push ecx
'00720cd7    8d55e4                  lea edx, dword ptr [ebp-1c]
'00720cda    52                      push edx
'00720cdb    8d45e8                  lea eax, dword ptr [ebp-18]
'00720cde    50                      push eax
'00720cdf    6a08                    push 08

' *** Reference to "__vbaFreeStrList"
'00720ce1    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4724, -4728, -4788, -4792, -4796, -4800, -4804)
'00720ce7    8d4db4                  lea ecx, dword ptr [ebp-4c]
'00720cea    51                      push ecx
'00720ceb    8d55b8                  lea edx, dword ptr [ebp-48]
'00720cee    52                      push edx
'00720cef    8d45bc                  lea eax, dword ptr [ebp-44]
'00720cf2    50                      push eax
'00720cf3    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00720cf6    51                      push ecx
'00720cf7    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'00720cf9    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58, var_61, var_62)
'00720cff    8b16                    mov edx, dword ptr [esi]
'00720d01    83c438                  add esp, 38
'00720d04    56                      push esi
'00720d05    ff9204030000            call dword ptr [edx+00000304]

' *** Reference to "__vbaObjSet"
'00720d0b    8b1db4104000            mov ebx, dword ptr [004010b4]
'00720d11    50                      push eax
'00720d12    8d45bc                  lea eax, dword ptr [ebp-44]
'00720d15    50                      push eax
'00720d16    ffd3                    call ebx
Set var_58 = 
'00720d18    8b0e                    mov ecx, dword ptr [esi]
'00720d1a    56                      push esi
'00720d1b    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'00720d21    ff9104030000            call dword ptr [ecx+00000304]
'00720d27    50                      push eax
'00720d28    8d55c0                  lea edx, dword ptr [ebp-40]
'00720d2b    52                      push edx
'00720d2c    ffd3                    call ebx
Set var_5 = var_58
'00720d2e    8d4de8                  lea ecx, dword ptr [ebp-18]
'00720d31    8bd8                    mov ebx, eax
'00720d33    8b03                    mov eax, dword ptr [ebx]
'00720d35    51                      push ecx
'00720d36    53                      push ebx
'00720d37    ff90a0000000            call dword ptr [eax+000000a0]
'00720d3d    dbe2                    fnclex
'00720d3f    85c0                    test ax, ax
'00720d41    7d12                    jge 720d55
'00720d43    68a0000000              push 000000a0
'00720d48    68d00d4300              push 00430dd0
'00720d4d    53                      push ebx
'00720d4e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720d4f    ff1580104000            call dword ptr [00401080]
'00720d55    8b45e8                  mov eax, dword ptr [ebp-18]
'00720d58    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'00720d5e    8b1a                    mov ebx, dword ptr [edx]
'00720d60    50                      push eax
'00720d61    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'00720d66    ff1570104000            call dword ptr [00401070]
var_49 = (vbNullString) & (vbCrLf)
'00720d6c    8bd0                    mov edx, eax
'00720d6e    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00720d71    ffd7                    call edi
'00720d73    50                      push eax
'00720d74    68b4734500              push 004573b4

' *** Reference to "__vbaStrCat"
'00720d79    ff1570104000            call dword ptr [00401070]
var_2441 = (var_49) & ("La dîme")
'00720d7f    8bd0                    mov edx, eax
'00720d81    8d4de0                  lea ecx, dword ptr [ebp-20]
'00720d84    ffd7                    call edi
'00720d86    50                      push eax
'00720d87    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'00720d8c    ff1570104000            call dword ptr [00401070]
var_98 = (var_2441) & (vbCrLf)
'00720d92    8bd0                    mov edx, eax
'00720d94    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00720d97    ffd7                    call edi
'00720d99    8bcb                    mov ecx, ebx
'00720d9b    8b9d20ffffff            mov ebx, dword ptr [ebp+ffffff20]
'00720da1    50                      push eax
'00720da2    53                      push ebx
'00720da3    ff91a4000000            call dword ptr [ecx+000000a4]
'00720da9    dbe2                    fnclex
'00720dab    85c0                    test ax, ax
'00720dad    7d12                    jge 720dc1
'00720daf    68a4000000              push 000000a4
'00720db4    68d00d4300              push 00430dd0
'00720db9    53                      push ebx
'00720dba    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720dbb    ff1580104000            call dword ptr [00401080]
'00720dc1    8d55dc                  lea edx, dword ptr [ebp-24]
'00720dc4    52                      push edx
'00720dc5    8d45e0                  lea eax, dword ptr [ebp-20]
'00720dc8    50                      push eax
'00720dc9    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00720dcc    51                      push ecx
'00720dcd    8d55e8                  lea edx, dword ptr [ebp-18]
'00720dd0    52                      push edx
'00720dd1    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'00720dd3    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4808, -4812, -4816)
'00720dd9    8d45bc                  lea eax, dword ptr [ebp-44]
'00720ddc    50                      push eax
'00720ddd    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00720de0    51                      push ecx
'00720de1    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00720de3    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58)
'00720de9    8b16                    mov edx, dword ptr [esi]
'00720deb    83c420                  add esp, 20
'00720dee    56                      push esi
'00720def    ff9218030000            call dword ptr [edx+00000318]
'00720df5    50                      push eax
'00720df6    8d45bc                  lea eax, dword ptr [ebp-44]
'00720df9    50                      push eax

' *** Reference to "__vbaObjSet"
'00720dfa    ff15b4104000            call dword ptr [004010b4]
Set var_58 = 
'00720e00    8d55e4                  lea edx, dword ptr [ebp-1c]
'00720e03    8bd8                    mov ebx, eax
'00720e05    8b0b                    mov ecx, dword ptr [ebx]
'00720e07    52                      push edx
'00720e08    53                      push ebx
'00720e09    ff91a0000000            call dword ptr [ecx+000000a0]
'00720e0f    dbe2                    fnclex
'00720e11    85c0                    test ax, ax
'00720e13    7d12                    jge 720e27
'00720e15    68a0000000              push 000000a0
'00720e1a    68d00d4300              push 00430dd0
'00720e1f    53                      push ebx
'00720e20    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720e21    ff1580104000            call dword ptr [00401080]
'00720e27    8b06                    mov eax, dword ptr [esi]
'00720e29    56                      push esi
'00720e2a    ff9008030000            call dword ptr [eax+00000308]
'00720e30    50                      push eax
'00720e31    8d4db8                  lea ecx, dword ptr [ebp-48]
'00720e34    51                      push ecx

' *** Reference to "__vbaObjSet"
'00720e35    ff15b4104000            call dword ptr [004010b4]
Set var_61 = Nothing
'00720e3b    8bd8                    mov ebx, eax
'00720e3d    8b13                    mov edx, dword ptr [ebx]
'00720e3f    8d45e0                  lea eax, dword ptr [ebp-20]
'00720e42    50                      push eax
'00720e43    53                      push ebx
'00720e44    ff92a0000000            call dword ptr [edx+000000a0]
'00720e4a    dbe2                    fnclex
'00720e4c    85c0                    test ax, ax
'00720e4e    7d12                    jge 720e62
'00720e50    68a0000000              push 000000a0
'00720e55    68d00d4300              push 00430dd0
'00720e5a    53                      push ebx
'00720e5b    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720e5c    ff1580104000            call dword ptr [00401080]
'00720e62    8b0e                    mov ecx, dword ptr [esi]
'00720e64    56                      push esi
'00720e65    ff9104030000            call dword ptr [ecx+00000304]

' *** Reference to "__vbaObjSet"
'00720e6b    8b1db4104000            mov ebx, dword ptr [004010b4]
'00720e71    50                      push eax
'00720e72    8d55b4                  lea edx, dword ptr [ebp-4c]
'00720e75    52                      push edx
'00720e76    ffd3                    call ebx
Set var_62 = 
'00720e78    89850cffffff            mov dword ptr [ebp+ffffff0c], eax
'00720e7e    8b06                    mov eax, dword ptr [esi]
'00720e80    56                      push esi
'00720e81    ff9004030000            call dword ptr [eax+00000304]
'00720e87    50                      push eax
'00720e88    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00720e8b    51                      push ecx
'00720e8c    ffd3                    call ebx
Set var_5 = Nothing
'00720e8e    8bd8                    mov ebx, eax
'00720e90    8b13                    mov edx, dword ptr [ebx]
'00720e92    8d45e8                  lea eax, dword ptr [ebp-18]
'00720e95    50                      push eax
'00720e96    53                      push ebx
'00720e97    ff92a0000000            call dword ptr [edx+000000a0]
'00720e9d    dbe2                    fnclex
'00720e9f    85c0                    test ax, ax
'00720ea1    7d12                    jge 720eb5
'00720ea3    68a0000000              push 000000a0
'00720ea8    68d00d4300              push 00430dd0
'00720ead    53                      push ebx
'00720eae    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720eaf    ff1580104000            call dword ptr [00401080]
'00720eb5    8b4de4                  mov ecx, dword ptr [ebp-1c]
'00720eb8    51                      push ecx

' *** Reference to "__vbaI4Str"
'00720eb9    ff1550124000            call dword ptr [00401250]
'00720ebf    8b16                    mov edx, dword ptr [esi]
'00720ec1    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'00720ec7    8d852cffffff            lea eax, dword ptr [ebp+ffffff2c]
'00720ecd    50                      push eax
'00720ece    8d8d30ffffff            lea ecx, dword ptr [ebp+ffffff30]
'00720ed4    51                      push ecx
'00720ed5    56                      push esi
'00720ed6    ff9208070000            call dword ptr [edx+00000708]
'00720edc    85c0                    test ax, ax
'00720ede    7d12                    jge 720ef2
'00720ee0    6808070000              push 00000708
'00720ee5    68b0154300              push 004315b0
'00720eea    56                      push esi
'00720eeb    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720eec    ff1580104000            call dword ptr [00401080]
'00720ef2    8b45e8                  mov eax, dword ptr [ebp-18]
'00720ef5    8b950cffffff            mov edx, dword ptr [ebp+ffffff0c]
'00720efb    8b1a                    mov ebx, dword ptr [edx]
'00720efd    50                      push eax
'00720efe    68c8734500              push 004573c8

' *** Reference to "__vbaStrCat"
'00720f03    ff1570104000            call dword ptr [00401070]
var_49 = (vbNullString) & ("Dîme en or : ")
'00720f09    8bd0                    mov edx, eax
'00720f0b    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00720f0e    ffd7                    call edi
'00720f10    8b4de0                  mov ecx, dword ptr [ebp-20]
'00720f13    50                      push eax
'00720f14    51                      push ecx

' *** Reference to "__vbaR8Str"
'00720f15    ff1524124000            call dword ptr [00401224]
'00720f1b    db852cffffff            fild dword ptr [ebp+ffffff2c]
'00720f21    83ec08                  sub esp, 08
'00720f24    dd9dacfeffff            fstp qword ptr [ebp+fffffeac]
'var_520 = (00)
'00720f2a    dc8dacfeffff            fmul qword ptr [ebp+fffffeac]
'00720f30    833d00b0720000          cmp dword ptr [0072b000], 00
'00720f37    7508                    jne 720f41

If (vbNullChar = 0) Then
'00720f39    dc35286c4000            fdiv qword ptr [00406c28]
'00720f3f    eb11                    jmp 720f52
    
End If
'00720f41    ff352c6c4000            push dword ptr [00406c2c]
'00720f47    ff35286c4000            push dword ptr [00406c28]

' *** Reference to "_adj_fdiv_m64"
'00720f4d    e83263ceff              call 407284
'00720f52    dfe0                    fnstsw ax
'00720f54    a80d                    test al, 0d
'00720f56    0f859a0c0000            jne 721bf6
'00720f5c    dd1c24                  fstp qword ptr [esp]
'var_62 = (00)

' *** Reference to "__vbaStrR8"
'00720f5f    ff1580114000            call dword ptr [00401180]
'00720f65    8bd0                    mov edx, eax
'00720f67    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00720f6a    ffd7                    call edi
'00720f6c    50                      push eax

' *** Reference to "__vbaStrCat"
'00720f6d    ff1570104000            call dword ptr [00401070]
var_2438 = (var_49) & (CStr(((CDbl(var_2441) * ()) / 100#)))
'00720f73    8bd0                    mov edx, eax
'00720f75    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00720f78    ffd7                    call edi
'00720f7a    50                      push eax
'00720f7b    68b4704500              push 004570b4

' *** Reference to "__vbaStrCat"
'00720f80    ff1570104000            call dword ptr [00401070]
var_1595 = (var_2438) & (" Po/mois")
'00720f86    8bd0                    mov edx, eax
'00720f88    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00720f8b    ffd7                    call edi
'00720f8d    50                      push eax
'00720f8e    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'00720f93    ff1570104000            call dword ptr [00401070]
var_1119 = (var_1595) & (vbCrLf)
'00720f99    8bd0                    mov edx, eax
'00720f9b    8d4dcc                  lea ecx, dword ptr [ebp-34]
'00720f9e    ffd7                    call edi
'00720fa0    8bd3                    mov edx, ebx
'00720fa2    8b9d0cffffff            mov ebx, dword ptr [ebp+ffffff0c]
'00720fa8    50                      push eax
'00720fa9    53                      push ebx
'00720faa    ff92a4000000            call dword ptr [edx+000000a4]
'00720fb0    dbe2                    fnclex
'00720fb2    85c0                    test ax, ax
'00720fb4    7d12                    jge 720fc8
'00720fb6    68a4000000              push 000000a4
'00720fbb    68d00d4300              push 00430dd0
'00720fc0    53                      push ebx
'00720fc1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00720fc2    ff1580104000            call dword ptr [00401080]
'00720fc8    8d45cc                  lea eax, dword ptr [ebp-34]
'00720fcb    50                      push eax
'00720fcc    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00720fcf    51                      push ecx
'00720fd0    8d55d4                  lea edx, dword ptr [ebp-2c]
'00720fd3    52                      push edx
'00720fd4    8d45d8                  lea eax, dword ptr [ebp-28]
'00720fd7    50                      push eax
'00720fd8    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00720fdb    51                      push ecx
'00720fdc    8d55e0                  lea edx, dword ptr [ebp-20]
'00720fdf    52                      push edx
'00720fe0    8d45e4                  lea eax, dword ptr [ebp-1c]
'00720fe3    50                      push eax
'00720fe4    8d4de8                  lea ecx, dword ptr [ebp-18]
'00720fe7    51                      push ecx
'00720fe8    6a08                    push 08

' *** Reference to "__vbaFreeStrList"
'00720fea    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4808, -4812, -4824, -4828, -4832, -4836, -4840)
'00720ff0    8d55b4                  lea edx, dword ptr [ebp-4c]
'00720ff3    52                      push edx
'00720ff4    8d45b8                  lea eax, dword ptr [ebp-48]
'00720ff7    50                      push eax
'00720ff8    8d4dbc                  lea ecx, dword ptr [ebp-44]
'00720ffb    51                      push ecx
'00720ffc    8d55c0                  lea edx, dword ptr [ebp-40]
'00720fff    52                      push edx
'00721000    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'00721002    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58, var_61, var_62)
'00721008    8b06                    mov eax, dword ptr [esi]
'0072100a    83c438                  add esp, 38
'0072100d    56                      push esi
'0072100e    ff9018030000            call dword ptr [eax+00000318]
'00721014    50                      push eax
'00721015    8d4dbc                  lea ecx, dword ptr [ebp-44]
'00721018    51                      push ecx

' *** Reference to "__vbaObjSet"
'00721019    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Nothing
'0072101f    8bd8                    mov ebx, eax
'00721021    8b13                    mov edx, dword ptr [ebx]
'00721023    8d45e4                  lea eax, dword ptr [ebp-1c]
'00721026    50                      push eax
'00721027    53                      push ebx
'00721028    ff92a0000000            call dword ptr [edx+000000a0]
'0072102e    dbe2                    fnclex
'00721030    85c0                    test ax, ax
'00721032    7d12                    jge 721046
'00721034    68a0000000              push 000000a0
'00721039    68d00d4300              push 00430dd0
'0072103e    53                      push ebx
'0072103f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00721040    ff1580104000            call dword ptr [00401080]
'00721046    8b0e                    mov ecx, dword ptr [esi]
'00721048    56                      push esi
'00721049    ff9108030000            call dword ptr [ecx+00000308]
'0072104f    50                      push eax
'00721050    8d55b8                  lea edx, dword ptr [ebp-48]
'00721053    52                      push edx

' *** Reference to "__vbaObjSet"
'00721054    ff15b4104000            call dword ptr [004010b4]
Set var_61 = 
'0072105a    8d4de0                  lea ecx, dword ptr [ebp-20]
'0072105d    8bd8                    mov ebx, eax
'0072105f    8b03                    mov eax, dword ptr [ebx]
'00721061    51                      push ecx
'00721062    53                      push ebx
'00721063    ff90a0000000            call dword ptr [eax+000000a0]
'00721069    dbe2                    fnclex
'0072106b    85c0                    test ax, ax
'0072106d    7d12                    jge 721081
'0072106f    68a0000000              push 000000a0
'00721074    68d00d4300              push 00430dd0
'00721079    53                      push ebx
'0072107a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072107b    ff1580104000            call dword ptr [00401080]
'00721081    8b16                    mov edx, dword ptr [esi]
'00721083    56                      push esi
'00721084    ff9204030000            call dword ptr [edx+00000304]

' *** Reference to "__vbaObjSet"
'0072108a    8b1db4104000            mov ebx, dword ptr [004010b4]
'00721090    50                      push eax
'00721091    8d45b4                  lea eax, dword ptr [ebp-4c]
'00721094    50                      push eax
'00721095    ffd3                    call ebx
Set var_62 = Nothing
'00721097    8b0e                    mov ecx, dword ptr [esi]
'00721099    56                      push esi
'0072109a    89850cffffff            mov dword ptr [ebp+ffffff0c], eax
'007210a0    ff9104030000            call dword ptr [ecx+00000304]
'007210a6    50                      push eax
'007210a7    8d55c0                  lea edx, dword ptr [ebp-40]
'007210aa    52                      push edx
'007210ab    ffd3                    call ebx
Set var_5 = Nothing
'007210ad    8d4de8                  lea ecx, dword ptr [ebp-18]
'007210b0    8bd8                    mov ebx, eax
'007210b2    8b03                    mov eax, dword ptr [ebx]
'007210b4    51                      push ecx
'007210b5    53                      push ebx
'007210b6    ff90a0000000            call dword ptr [eax+000000a0]
'007210bc    dbe2                    fnclex
'007210be    85c0                    test ax, ax
'007210c0    7d12                    jge 7210d4
'007210c2    68a0000000              push 000000a0
'007210c7    68d00d4300              push 00430dd0
'007210cc    53                      push ebx
'007210cd    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007210ce    ff1580104000            call dword ptr [00401080]
'007210d4    8b55e4                  mov edx, dword ptr [ebp-1c]
'007210d7    52                      push edx

' *** Reference to "__vbaI4Str"
'007210d8    ff1550124000            call dword ptr [00401250]
'007210de    8d8d2cffffff            lea ecx, dword ptr [ebp+ffffff2c]
'007210e4    51                      push ecx
'007210e5    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'007210eb    52                      push edx
'007210ec    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'007210f2    8b06                    mov eax, dword ptr [esi]
'007210f4    56                      push esi
'007210f5    ff900c070000            call dword ptr [eax+0000070c]
'007210fb    85c0                    test ax, ax
'007210fd    7d12                    jge 721111
'007210ff    680c070000              push 0000070c
'00721104    68b0154300              push 004315b0
'00721109    56                      push esi
'0072110a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072110b    ff1580104000            call dword ptr [00401080]
'00721111    8b4de8                  mov ecx, dword ptr [ebp-18]
'00721114    8b850cffffff            mov eax, dword ptr [ebp+ffffff0c]
'0072111a    8b18                    mov ebx, dword ptr [eax]
'0072111c    51                      push ecx
'0072111d    68e8734500              push 004573e8

' *** Reference to "__vbaStrCat"
'00721122    ff1570104000            call dword ptr [00401070]
var_49 = () & ("Dîme en matières premières : ")
'00721128    8bd0                    mov edx, eax
'0072112a    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0072112d    ffd7                    call edi
'0072112f    8b55e0                  mov edx, dword ptr [ebp-20]
'00721132    50                      push eax
'00721133    52                      push edx

' *** Reference to "__vbaR8Str"
'00721134    ff1524124000            call dword ptr [00401224]
'0072113a    db852cffffff            fild dword ptr [ebp+ffffff2c]
'00721140    83ec08                  sub esp, 08
'00721143    dd9da0feffff            fstp qword ptr [ebp+fffffea0]
'var_527 = (00)
'00721149    dc8da0feffff            fmul qword ptr [ebp+fffffea0]
'0072114f    833d00b0720000          cmp dword ptr [0072b000], 00
'00721156    7508                    jne 721160

If (vbNullChar = 0) Then
'00721158    dc35286c4000            fdiv qword ptr [00406c28]
'0072115e    eb11                    jmp 721171
    
End If
'00721160    ff352c6c4000            push dword ptr [00406c2c]
'00721166    ff35286c4000            push dword ptr [00406c28]

' *** Reference to "_adj_fdiv_m64"
'0072116c    e81361ceff              call 407284
'00721171    dfe0                    fnstsw ax
'00721173    a80d                    test al, 0d
'00721175    0f857b0a0000            jne 721bf6
'0072117b    dd1c24                  fstp qword ptr [esp]
'var_44 = (00)

' *** Reference to "__vbaStrR8"
'0072117e    ff1580114000            call dword ptr [00401180]
'00721184    8bd0                    mov edx, eax
'00721186    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00721189    ffd7                    call edi
'0072118b    50                      push eax

' *** Reference to "__vbaStrCat"
'0072118c    ff1570104000            call dword ptr [00401070]
var_2396 = (var_49) & (CStr(((CDbl(CLng()) * ()) / 100#)))
'00721192    8bd0                    mov edx, eax
'00721194    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00721197    ffd7                    call edi
'00721199    50                      push eax
'0072119a    68b4704500              push 004570b4

' *** Reference to "__vbaStrCat"
'0072119f    ff1570104000            call dword ptr [00401070]
var_101 = (var_2396) & (" Po/mois")
'007211a5    8bd0                    mov edx, eax
'007211a7    8d4dd0                  lea ecx, dword ptr [ebp-30]
'007211aa    ffd7                    call edi
'007211ac    50                      push eax
'007211ad    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'007211b2    ff1570104000            call dword ptr [00401070]
var_1632 = (var_101) & (vbCrLf)
'007211b8    8bd0                    mov edx, eax
'007211ba    8d4dcc                  lea ecx, dword ptr [ebp-34]
'007211bd    ffd7                    call edi
'007211bf    50                      push eax
'007211c0    8bc3                    mov eax, ebx
'007211c2    8b9d0cffffff            mov ebx, dword ptr [ebp+ffffff0c]
'007211c8    53                      push ebx
'007211c9    ff90a4000000            call dword ptr [eax+000000a4]
'007211cf    dbe2                    fnclex
'007211d1    85c0                    test ax, ax
'007211d3    7d12                    jge 7211e7
'007211d5    68a4000000              push 000000a4
'007211da    68d00d4300              push 00430dd0
'007211df    53                      push ebx
'007211e0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007211e1    ff1580104000            call dword ptr [00401080]
'007211e7    8d4dcc                  lea ecx, dword ptr [ebp-34]
'007211ea    51                      push ecx
'007211eb    8d55d0                  lea edx, dword ptr [ebp-30]
'007211ee    52                      push edx
'007211ef    8d45d4                  lea eax, dword ptr [ebp-2c]
'007211f2    50                      push eax
'007211f3    8d4dd8                  lea ecx, dword ptr [ebp-28]
'007211f6    51                      push ecx
'007211f7    8d55dc                  lea edx, dword ptr [ebp-24]
'007211fa    52                      push edx
'007211fb    8d45e0                  lea eax, dword ptr [ebp-20]
'007211fe    50                      push eax
'007211ff    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00721202    51                      push ecx
'00721203    8d55e8                  lea edx, dword ptr [ebp-18]
'00721206    52                      push edx
'00721207    6a08                    push 08

' *** Reference to "__vbaFreeStrList"
'00721209    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , , , -4864)
'0072120f    8d45b4                  lea eax, dword ptr [ebp-4c]
'00721212    50                      push eax
'00721213    8d4db8                  lea ecx, dword ptr [ebp-48]
'00721216    51                      push ecx
'00721217    8d55bc                  lea edx, dword ptr [ebp-44]
'0072121a    52                      push edx
'0072121b    8d45c0                  lea eax, dword ptr [ebp-40]
'0072121e    50                      push eax
'0072121f    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'00721221    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58, var_61, var_62)
'00721227    8b0e                    mov ecx, dword ptr [esi]
'00721229    83c438                  add esp, 38
'0072122c    56                      push esi
'0072122d    ff9118030000            call dword ptr [ecx+00000318]
'00721233    50                      push eax
'00721234    8d55bc                  lea edx, dword ptr [ebp-44]
'00721237    52                      push edx

' *** Reference to "__vbaObjSet"
'00721238    ff15b4104000            call dword ptr [004010b4]
Set var_58 = 
'0072123e    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00721241    8bd8                    mov ebx, eax
'00721243    8b03                    mov eax, dword ptr [ebx]
'00721245    51                      push ecx
'00721246    53                      push ebx
'00721247    ff90a0000000            call dword ptr [eax+000000a0]
'0072124d    dbe2                    fnclex
'0072124f    85c0                    test ax, ax
'00721251    7d12                    jge 721265
'00721253    68a0000000              push 000000a0
'00721258    68d00d4300              push 00430dd0
'0072125d    53                      push ebx
'0072125e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072125f    ff1580104000            call dword ptr [00401080]
'00721265    8b16                    mov edx, dword ptr [esi]
'00721267    56                      push esi
'00721268    ff9208030000            call dword ptr [edx+00000308]
'0072126e    50                      push eax
'0072126f    8d45b8                  lea eax, dword ptr [ebp-48]
'00721272    50                      push eax

' *** Reference to "__vbaObjSet"
'00721273    ff15b4104000            call dword ptr [004010b4]
Set var_61 = -284
'00721279    8d55e0                  lea edx, dword ptr [ebp-20]
'0072127c    8bd8                    mov ebx, eax
'0072127e    8b0b                    mov ecx, dword ptr [ebx]
'00721280    52                      push edx
'00721281    53                      push ebx
'00721282    ff91a0000000            call dword ptr [ecx+000000a0]
'00721288    dbe2                    fnclex
'0072128a    85c0                    test ax, ax
'0072128c    7d12                    jge 7212a0
'0072128e    68a0000000              push 000000a0
'00721293    68d00d4300              push 00430dd0
'00721298    53                      push ebx
'00721299    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072129a    ff1580104000            call dword ptr [00401080]
'007212a0    8b06                    mov eax, dword ptr [esi]
'007212a2    56                      push esi
'007212a3    ff9004030000            call dword ptr [eax+00000304]

' *** Reference to "__vbaObjSet"
'007212a9    8b1db4104000            mov ebx, dword ptr [004010b4]
'007212af    50                      push eax
'007212b0    8d4db4                  lea ecx, dword ptr [ebp-4c]
'007212b3    51                      push ecx
'007212b4    ffd3                    call ebx
Set var_62 = Nothing
'007212b6    8b16                    mov edx, dword ptr [esi]
'007212b8    56                      push esi
'007212b9    89850cffffff            mov dword ptr [ebp+ffffff0c], eax
'007212bf    ff9204030000            call dword ptr [edx+00000304]
'007212c5    50                      push eax
'007212c6    8d45c0                  lea eax, dword ptr [ebp-40]
'007212c9    50                      push eax
'007212ca    ffd3                    call ebx
Set var_5 = Nothing
'007212cc    8d55e8                  lea edx, dword ptr [ebp-18]
'007212cf    8bd8                    mov ebx, eax
'007212d1    8b0b                    mov ecx, dword ptr [ebx]
'007212d3    52                      push edx
'007212d4    53                      push ebx
'007212d5    ff91a0000000            call dword ptr [ecx+000000a0]
'007212db    dbe2                    fnclex
'007212dd    85c0                    test ax, ax
'007212df    7d12                    jge 7212f3
'007212e1    68a0000000              push 000000a0
'007212e6    68d00d4300              push 00430dd0
'007212eb    53                      push ebx
'007212ec    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007212ed    ff1580104000            call dword ptr [00401080]
'007212f3    8b45e4                  mov eax, dword ptr [ebp-1c]
'007212f6    50                      push eax

' *** Reference to "__vbaI4Str"
'007212f7    ff1550124000            call dword ptr [00401250]
'007212fd    8b0e                    mov ecx, dword ptr [esi]
'007212ff    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'00721305    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'0072130b    52                      push edx
'0072130c    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'00721312    50                      push eax
'00721313    56                      push esi
'00721314    ff9104070000            call dword ptr [ecx+00000704]
'0072131a    85c0                    test ax, ax
'0072131c    7d12                    jge 721330
'0072131e    6804070000              push 00000704
'00721323    68b0154300              push 004315b0
'00721328    56                      push esi
'00721329    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072132a    ff1580104000            call dword ptr [00401080]
'00721330    8b55e8                  mov edx, dword ptr [ebp-18]
'00721333    8b8d0cffffff            mov ecx, dword ptr [ebp+ffffff0c]
'00721339    8b19                    mov ebx, dword ptr [ecx]
'0072133b    52                      push edx
'0072133c    6828744500              push 00457428

' *** Reference to "__vbaStrCat"
'00721341    ff1570104000            call dword ptr [00401070]
var_23 = () & ("Dîme totaux : ")
'00721347    8bd0                    mov edx, eax
'00721349    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0072134c    ffd7                    call edi
'0072134e    50                      push eax
'0072134f    8b45e0                  mov eax, dword ptr [ebp-20]
'00721352    50                      push eax

' *** Reference to "__vbaR8Str"
'00721353    ff1524124000            call dword ptr [00401224]
'00721359    db852cffffff            fild dword ptr [ebp+ffffff2c]
'0072135f    83ec08                  sub esp, 08
'00721362    dd9d94feffff            fstp qword ptr [ebp+fffffe94]
'var_538 = (00)
'00721368    dc8d94feffff            fmul qword ptr [ebp+fffffe94]
'0072136e    833d00b0720000          cmp dword ptr [0072b000], 00
'00721375    7508                    jne 72137f

If (vbNullChar = 0) Then
'00721377    dc35286c4000            fdiv qword ptr [00406c28]
'0072137d    eb11                    jmp 721390
    
End If
'0072137f    ff352c6c4000            push dword ptr [00406c2c]
'00721385    ff35286c4000            push dword ptr [00406c28]

' *** Reference to "_adj_fdiv_m64"
'0072138b    e8f45eceff              call 407284
'00721390    dfe0                    fnstsw ax
'00721392    a80d                    test al, 0d
'00721394    0f855c080000            jne 721bf6
'0072139a    dd1c24                  fstp qword ptr [esp]
'var_2101 = (00)

' *** Reference to "__vbaStrR8"
'0072139d    ff1580114000            call dword ptr [00401180]
'007213a3    8bd0                    mov edx, eax
'007213a5    8d4dd8                  lea ecx, dword ptr [ebp-28]
'007213a8    ffd7                    call edi
'007213aa    50                      push eax

' *** Reference to "__vbaStrCat"
'007213ab    ff1570104000            call dword ptr [00401070]
var_2435 = (var_23) & (CStr(((CDbl(CLng()) * ()) / 100#)))
'007213b1    8bd0                    mov edx, eax
'007213b3    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'007213b6    ffd7                    call edi
'007213b8    50                      push eax
'007213b9    68b4704500              push 004570b4

' *** Reference to "__vbaStrCat"
'007213be    ff1570104000            call dword ptr [00401070]
var_2344 = (var_2435) & (" Po/mois")
'007213c4    8bd0                    mov edx, eax
'007213c6    8d4dd0                  lea ecx, dword ptr [ebp-30]
'007213c9    ffd7                    call edi
'007213cb    50                      push eax
'007213cc    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'007213d1    ff1570104000            call dword ptr [00401070]
var_1647 = (var_2344) & (vbCrLf)
'007213d7    8bd0                    mov edx, eax
'007213d9    8d4dcc                  lea ecx, dword ptr [ebp-34]
'007213dc    ffd7                    call edi
'007213de    8bcb                    mov ecx, ebx
'007213e0    8b9d0cffffff            mov ebx, dword ptr [ebp+ffffff0c]
'007213e6    50                      push eax
'007213e7    53                      push ebx
'007213e8    ff91a4000000            call dword ptr [ecx+000000a4]
'007213ee    dbe2                    fnclex
'007213f0    85c0                    test ax, ax
'007213f2    7d12                    jge 721406
'007213f4    68a4000000              push 000000a4
'007213f9    68d00d4300              push 00430dd0
'007213fe    53                      push ebx
'007213ff    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00721400    ff1580104000            call dword ptr [00401080]
'00721406    8d55cc                  lea edx, dword ptr [ebp-34]
'00721409    52                      push edx
'0072140a    8d45d0                  lea eax, dword ptr [ebp-30]
'0072140d    50                      push eax
'0072140e    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00721411    51                      push ecx
'00721412    8d55d8                  lea edx, dword ptr [ebp-28]
'00721415    52                      push edx
'00721416    8d45dc                  lea eax, dword ptr [ebp-24]
'00721419    50                      push eax
'0072141a    8d4de0                  lea ecx, dword ptr [ebp-20]
'0072141d    51                      push ecx
'0072141e    8d55e4                  lea edx, dword ptr [ebp-1c]
'00721421    52                      push edx
'00721422    8d45e8                  lea eax, dword ptr [ebp-18]
'00721425    50                      push eax
'00721426    6a08                    push 08

' *** Reference to "__vbaFreeStrList"
'00721428    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , , , , , -4880, -4884, -4888)
'0072142e    8d4db4                  lea ecx, dword ptr [ebp-4c]
'00721431    51                      push ecx
'00721432    8d55b8                  lea edx, dword ptr [ebp-48]
'00721435    52                      push edx
'00721436    8d45bc                  lea eax, dword ptr [ebp-44]
'00721439    50                      push eax
'0072143a    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0072143d    51                      push ecx
'0072143e    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'00721440    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58, var_61, var_62)
'00721446    8b16                    mov edx, dword ptr [esi]
'00721448    83c438                  add esp, 38
'0072144b    56                      push esi
'0072144c    ff9204030000            call dword ptr [edx+00000304]

' *** Reference to "__vbaObjSet"
'00721452    8b1db4104000            mov ebx, dword ptr [004010b4]
'00721458    50                      push eax
'00721459    8d45bc                  lea eax, dword ptr [ebp-44]
'0072145c    50                      push eax
'0072145d    ffd3                    call ebx
Set var_58 = 
'0072145f    8b0e                    mov ecx, dword ptr [esi]
'00721461    56                      push esi
'00721462    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'00721468    ff9104030000            call dword ptr [ecx+00000304]
'0072146e    50                      push eax
'0072146f    8d55c0                  lea edx, dword ptr [ebp-40]
'00721472    52                      push edx
'00721473    ffd3                    call ebx
Set var_5 = var_58
'00721475    8d4de8                  lea ecx, dword ptr [ebp-18]
'00721478    8bd8                    mov ebx, eax
'0072147a    8b03                    mov eax, dword ptr [ebx]
'0072147c    51                      push ecx
'0072147d    53                      push ebx
'0072147e    ff90a0000000            call dword ptr [eax+000000a0]
'00721484    dbe2                    fnclex
'00721486    85c0                    test ax, ax
'00721488    7d12                    jge 72149c
'0072148a    68a0000000              push 000000a0
'0072148f    68d00d4300              push 00430dd0
'00721494    53                      push ebx
'00721495    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00721496    ff1580104000            call dword ptr [00401080]
'0072149c    8b45e8                  mov eax, dword ptr [ebp-18]
'0072149f    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'007214a5    8b1a                    mov ebx, dword ptr [edx]
'007214a7    50                      push eax
'007214a8    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'007214ad    ff1570104000            call dword ptr [00401070]
var_1473 = () & (vbCrLf)
'007214b3    8bd0                    mov edx, eax
'007214b5    8d4de4                  lea ecx, dword ptr [ebp-1c]
'007214b8    ffd7                    call edi
'007214ba    50                      push eax
'007214bb    684c744500              push 0045744c

' *** Reference to "__vbaStrCat"
'007214c0    ff1570104000            call dword ptr [00401070]
var_2114 = (var_1473) & ("Les instances")
'007214c6    8bd0                    mov edx, eax
'007214c8    8d4de0                  lea ecx, dword ptr [ebp-20]
'007214cb    ffd7                    call edi
'007214cd    50                      push eax
'007214ce    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'007214d3    ff1570104000            call dword ptr [00401070]
var_2395 = (var_2114) & (vbCrLf)
'007214d9    8bd0                    mov edx, eax
'007214db    8d4ddc                  lea ecx, dword ptr [ebp-24]
'007214de    ffd7                    call edi
'007214e0    8bcb                    mov ecx, ebx
'007214e2    8b9d20ffffff            mov ebx, dword ptr [ebp+ffffff20]
'007214e8    50                      push eax
'007214e9    53                      push ebx
'007214ea    ff91a4000000            call dword ptr [ecx+000000a4]
'007214f0    dbe2                    fnclex
'007214f2    85c0                    test ax, ax
'007214f4    7d12                    jge 721508
'007214f6    68a4000000              push 000000a4
'007214fb    68d00d4300              push 00430dd0
'00721500    53                      push ebx
'00721501    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00721502    ff1580104000            call dword ptr [00401080]
'00721508    8d55dc                  lea edx, dword ptr [ebp-24]
'0072150b    52                      push edx
'0072150c    8d45e0                  lea eax, dword ptr [ebp-20]
'0072150f    50                      push eax
'00721510    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00721513    51                      push ecx
'00721514    8d55e8                  lea edx, dword ptr [ebp-18]
'00721517    52                      push edx
'00721518    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'0072151a    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , -4892, -4896, -4900)
'00721520    8d45bc                  lea eax, dword ptr [ebp-44]
'00721523    50                      push eax
'00721524    8d4dc0                  lea ecx, dword ptr [ebp-40]
'00721527    51                      push ecx
'00721528    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0072152a    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58)
'00721530    8b16                    mov edx, dword ptr [esi]
'00721532    83c420                  add esp, 20
'00721535    56                      push esi
'00721536    ff9204030000            call dword ptr [edx+00000304]

' *** Reference to "__vbaObjSet"
'0072153c    8b1db4104000            mov ebx, dword ptr [004010b4]
'00721542    50                      push eax
'00721543    8d45b8                  lea eax, dword ptr [ebp-48]
'00721546    50                      push eax
'00721547    ffd3                    call ebx
Set var_61 = 
'00721549    8b0e                    mov ecx, dword ptr [esi]
'0072154b    56                      push esi
'0072154c    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'00721552    ff9104030000            call dword ptr [ecx+00000304]
'00721558    50                      push eax
'00721559    8d55bc                  lea edx, dword ptr [ebp-44]
'0072155c    52                      push edx
'0072155d    ffd3                    call ebx
Set var_58 = var_61
'0072155f    8d4de0                  lea ecx, dword ptr [ebp-20]
'00721562    8bd8                    mov ebx, eax
'00721564    8b03                    mov eax, dword ptr [ebx]
'00721566    51                      push ecx
'00721567    53                      push ebx
'00721568    ff90a0000000            call dword ptr [eax+000000a0]
'0072156e    dbe2                    fnclex
'00721570    85c0                    test ax, ax
'00721572    7d12                    jge 721586
'00721574    68a0000000              push 000000a0
'00721579    68d00d4300              push 00430dd0
'0072157e    53                      push ebx
'0072157f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00721580    ff1580104000            call dword ptr [00401080]
'00721586    8b16                    mov edx, dword ptr [esi]
'00721588    56                      push esi
'00721589    ff9218030000            call dword ptr [edx+00000318]
'0072158f    50                      push eax
'00721590    8d45c0                  lea eax, dword ptr [ebp-40]
'00721593    50                      push eax

' *** Reference to "__vbaObjSet"
'00721594    ff15b4104000            call dword ptr [004010b4]
Set var_5 = var_58
'0072159a    8d55e8                  lea edx, dword ptr [ebp-18]
'0072159d    8bd8                    mov ebx, eax
'0072159f    8b0b                    mov ecx, dword ptr [ebx]
'007215a1    52                      push edx
'007215a2    53                      push ebx
'007215a3    ff91a0000000            call dword ptr [ecx+000000a0]
'007215a9    dbe2                    fnclex
'007215ab    85c0                    test ax, ax
'007215ad    7d12                    jge 7215c1
'007215af    68a0000000              push 000000a0
'007215b4    68d00d4300              push 00430dd0
'007215b9    53                      push ebx
'007215ba    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007215bb    ff1580104000            call dword ptr [00401080]
'007215c1    8b45e8                  mov eax, dword ptr [ebp-18]
'007215c4    50                      push eax

' *** Reference to "__vbaI4Str"
'007215c5    ff1550124000            call dword ptr [00401250]
'007215cb    8b0e                    mov ecx, dword ptr [esi]
'007215cd    8d55e4                  lea edx, dword ptr [ebp-1c]
'007215d0    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'007215d6    52                      push edx
'007215d7    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'007215dd    50                      push eax
'007215de    56                      push esi
'007215df    ff9110070000            call dword ptr [ecx+00000710]
'007215e5    85c0                    test ax, ax
'007215e7    7d12                    jge 7215fb
'007215e9    6810070000              push 00000710
'007215ee    68b0154300              push 004315b0
'007215f3    56                      push esi
'007215f4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007215f5    ff1580104000            call dword ptr [00401080]
'007215fb    8b55e0                  mov edx, dword ptr [ebp-20]
'007215fe    8b45e4                  mov eax, dword ptr [ebp-1c]
'00721601    8b8d14ffffff            mov ecx, dword ptr [ebp+ffffff14]
'00721607    8b19                    mov ebx, dword ptr [ecx]
'00721609    52                      push edx
'0072160a    50                      push eax

' *** Reference to "__vbaStrCat"
'0072160b    ff1570104000            call dword ptr [00401070]
var_2114 = (var_2114) & (var_1473)
'00721611    8bd0                    mov edx, eax
'00721613    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00721616    ffd7                    call edi
'00721618    50                      push eax
'00721619    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'0072161e    ff1570104000            call dword ptr [00401070]
var_2342 = (var_2114) & (vbCrLf)
'00721624    8bd0                    mov edx, eax
'00721626    8d4dd8                  lea ecx, dword ptr [ebp-28]
'00721629    ffd7                    call edi
'0072162b    8bcb                    mov ecx, ebx
'0072162d    8b9d14ffffff            mov ebx, dword ptr [ebp+ffffff14]
'00721633    50                      push eax
'00721634    53                      push ebx
'00721635    ff91a4000000            call dword ptr [ecx+000000a4]
'0072163b    dbe2                    fnclex
'0072163d    85c0                    test ax, ax
'0072163f    7d12                    jge 721653
'00721641    68a4000000              push 000000a4
'00721646    68d00d4300              push 00430dd0
'0072164b    53                      push ebx
'0072164c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072164d    ff1580104000            call dword ptr [00401080]
'00721653    8d55d8                  lea edx, dword ptr [ebp-28]
'00721656    52                      push edx
'00721657    8d45dc                  lea eax, dword ptr [ebp-24]
'0072165a    50                      push eax
'0072165b    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0072165e    51                      push ecx
'0072165f    8d55e0                  lea edx, dword ptr [ebp-20]
'00721662    52                      push edx
'00721663    8d45e8                  lea eax, dword ptr [ebp-18]
'00721666    50                      push eax
'00721667    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'00721669    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , -4896, -4892, -4908, -4912)
'0072166f    8d4db8                  lea ecx, dword ptr [ebp-48]
'00721672    51                      push ecx
'00721673    8d55bc                  lea edx, dword ptr [ebp-44]
'00721676    52                      push edx
'00721677    8d45c0                  lea eax, dword ptr [ebp-40]
'0072167a    50                      push eax
'0072167b    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0072167d    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58, var_61)
'00721683    8b0e                    mov ecx, dword ptr [esi]
'00721685    83c428                  add esp, 28
'00721688    56                      push esi
'00721689    ff9104030000            call dword ptr [ecx+00000304]

' *** Reference to "__vbaObjSet"
'0072168f    8b1db4104000            mov ebx, dword ptr [004010b4]
'00721695    50                      push eax
'00721696    8d55bc                  lea edx, dword ptr [ebp-44]
'00721699    52                      push edx
'0072169a    ffd3                    call ebx
Set var_58 = 
'0072169c    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'007216a2    8b06                    mov eax, dword ptr [esi]
'007216a4    56                      push esi
'007216a5    ff9004030000            call dword ptr [eax+00000304]
'007216ab    50                      push eax
'007216ac    8d4dc0                  lea ecx, dword ptr [ebp-40]
'007216af    51                      push ecx
'007216b0    ffd3                    call ebx
Set var_5 = Nothing
'007216b2    8bd8                    mov ebx, eax
'007216b4    8b13                    mov edx, dword ptr [ebx]
'007216b6    8d45e8                  lea eax, dword ptr [ebp-18]
'007216b9    50                      push eax
'007216ba    53                      push ebx
'007216bb    ff92a0000000            call dword ptr [edx+000000a0]
'007216c1    dbe2                    fnclex
'007216c3    85c0                    test ax, ax
'007216c5    7d12                    jge 7216d9
'007216c7    68a0000000              push 000000a0
'007216cc    68d00d4300              push 00430dd0
'007216d1    53                      push ebx
'007216d2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007216d3    ff1580104000            call dword ptr [00401080]
'007216d9    8b55e8                  mov edx, dword ptr [ebp-18]
'007216dc    8b8d20ffffff            mov ecx, dword ptr [ebp+ffffff20]
'007216e2    8b19                    mov ebx, dword ptr [ecx]
'007216e4    52                      push edx
'007216e5    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'007216ea    ff1570104000            call dword ptr [00401070]
var_41 = () & (vbCrLf)
'007216f0    8bd0                    mov edx, eax
'007216f2    8d4de4                  lea ecx, dword ptr [ebp-1c]
'007216f5    ffd7                    call edi
'007216f7    50                      push eax
'007216f8    68a06f4500              push 00456fa0

' *** Reference to "__vbaStrCat"
'007216fd    ff1570104000            call dword ptr [00401070]
var_2282 = (var_41) & ("Les Pnjs")
'00721703    8bd0                    mov edx, eax
'00721705    8d4de0                  lea ecx, dword ptr [ebp-20]
'00721708    ffd7                    call edi
'0072170a    50                      push eax
'0072170b    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'00721710    ff1570104000            call dword ptr [00401070]
var_850 = (var_2282) & (vbCrLf)
'00721716    8bd0                    mov edx, eax
'00721718    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0072171b    ffd7                    call edi
'0072171d    50                      push eax
'0072171e    8bc3                    mov eax, ebx
'00721720    8b9d20ffffff            mov ebx, dword ptr [ebp+ffffff20]
'00721726    53                      push ebx
'00721727    ff90a4000000            call dword ptr [eax+000000a4]
'0072172d    dbe2                    fnclex
'0072172f    85c0                    test ax, ax
'00721731    7d12                    jge 721745
'00721733    68a4000000              push 000000a4
'00721738    68d00d4300              push 00430dd0
'0072173d    53                      push ebx
'0072173e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072173f    ff1580104000            call dword ptr [00401080]
'00721745    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00721748    51                      push ecx
'00721749    8d55e0                  lea edx, dword ptr [ebp-20]
'0072174c    52                      push edx
'0072174d    8d45e4                  lea eax, dword ptr [ebp-1c]
'00721750    50                      push eax
'00721751    8d4de8                  lea ecx, dword ptr [ebp-18]
'00721754    51                      push ecx
'00721755    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'00721757    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , -4916, -4920, -4924)
'0072175d    8d55bc                  lea edx, dword ptr [ebp-44]
'00721760    52                      push edx
'00721761    8d45c0                  lea eax, dword ptr [ebp-40]
'00721764    50                      push eax
'00721765    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00721767    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58)
'0072176d    8b0e                    mov ecx, dword ptr [esi]
'0072176f    83c420                  add esp, 20
'00721772    56                      push esi
'00721773    ff9104030000            call dword ptr [ecx+00000304]

' *** Reference to "__vbaObjSet"
'00721779    8b1db4104000            mov ebx, dword ptr [004010b4]
'0072177f    50                      push eax
'00721780    8d55b8                  lea edx, dword ptr [ebp-48]
'00721783    52                      push edx
'00721784    ffd3                    call ebx
Set var_61 = 
'00721786    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'0072178c    8b06                    mov eax, dword ptr [esi]
'0072178e    56                      push esi
'0072178f    ff9004030000            call dword ptr [eax+00000304]
'00721795    50                      push eax
'00721796    8d4dbc                  lea ecx, dword ptr [ebp-44]
'00721799    51                      push ecx
'0072179a    ffd3                    call ebx
Set var_58 = Nothing
'0072179c    8bd8                    mov ebx, eax
'0072179e    8b13                    mov edx, dword ptr [ebx]
'007217a0    8d45e0                  lea eax, dword ptr [ebp-20]
'007217a3    50                      push eax
'007217a4    53                      push ebx
'007217a5    ff92a0000000            call dword ptr [edx+000000a0]
'007217ab    dbe2                    fnclex
'007217ad    85c0                    test ax, ax
'007217af    7d12                    jge 7217c3
'007217b1    68a0000000              push 000000a0
'007217b6    68d00d4300              push 00430dd0
'007217bb    53                      push ebx
'007217bc    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007217bd    ff1580104000            call dword ptr [00401080]
'007217c3    8b0e                    mov ecx, dword ptr [esi]
'007217c5    56                      push esi
'007217c6    ff9118030000            call dword ptr [ecx+00000318]
'007217cc    50                      push eax
'007217cd    8d55c0                  lea edx, dword ptr [ebp-40]
'007217d0    52                      push edx

' *** Reference to "__vbaObjSet"
'007217d1    ff15b4104000            call dword ptr [004010b4]
Set var_5 = 
'007217d7    8d4de8                  lea ecx, dword ptr [ebp-18]
'007217da    8bd8                    mov ebx, eax
'007217dc    8b03                    mov eax, dword ptr [ebx]
'007217de    51                      push ecx
'007217df    53                      push ebx
'007217e0    ff90a0000000            call dword ptr [eax+000000a0]
'007217e6    dbe2                    fnclex
'007217e8    85c0                    test ax, ax
'007217ea    7d12                    jge 7217fe
'007217ec    68a0000000              push 000000a0
'007217f1    68d00d4300              push 00430dd0
'007217f6    53                      push ebx
'007217f7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007217f8    ff1580104000            call dword ptr [00401080]
'007217fe    8b55e8                  mov edx, dword ptr [ebp-18]
'00721801    52                      push edx

' *** Reference to "__vbaI4Str"
'00721802    ff1550124000            call dword ptr [00401250]
'00721808    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0072180b    51                      push ecx
'0072180c    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'00721812    52                      push edx
'00721813    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'00721819    8b06                    mov eax, dword ptr [esi]
'0072181b    56                      push esi
'0072181c    ff901c070000            call dword ptr [eax+0000071c]
'00721822    85c0                    test ax, ax
'00721824    7d12                    jge 721838
'00721826    681c070000              push 0000071c
'0072182b    68b0154300              push 004315b0
'00721830    56                      push esi
'00721831    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00721832    ff1580104000            call dword ptr [00401080]
'00721838    8b4de0                  mov ecx, dword ptr [ebp-20]
'0072183b    8b55e4                  mov edx, dword ptr [ebp-1c]
'0072183e    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'00721844    8b18                    mov ebx, dword ptr [eax]
'00721846    51                      push ecx
'00721847    52                      push edx

' *** Reference to "__vbaStrCat"
'00721848    ff1570104000            call dword ptr [00401070]
var_5 = (var_2282) & (var_41)
'0072184e    8bd0                    mov edx, eax
'00721850    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00721853    ffd7                    call edi
'00721855    50                      push eax
'00721856    8bc3                    mov eax, ebx
'00721858    8b9d14ffffff            mov ebx, dword ptr [ebp+ffffff14]
'0072185e    53                      push ebx
'0072185f    ff90a4000000            call dword ptr [eax+000000a4]
'00721865    dbe2                    fnclex
'00721867    85c0                    test ax, ax
'00721869    7d12                    jge 72187d
'0072186b    68a4000000              push 000000a4
'00721870    68d00d4300              push 00430dd0
'00721875    53                      push ebx
'00721876    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00721877    ff1580104000            call dword ptr [00401080]
'0072187d    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00721880    51                      push ecx
'00721881    8d55e4                  lea edx, dword ptr [ebp-1c]
'00721884    52                      push edx
'00721885    8d45e0                  lea eax, dword ptr [ebp-20]
'00721888    50                      push eax
'00721889    8d4de8                  lea ecx, dword ptr [ebp-18]
'0072188c    51                      push ecx
'0072188d    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'0072188f    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , -4920, -4916, -4932)
'00721895    8d55b8                  lea edx, dword ptr [ebp-48]
'00721898    52                      push edx
'00721899    8d45bc                  lea eax, dword ptr [ebp-44]
'0072189c    50                      push eax
'0072189d    8d4dc0                  lea ecx, dword ptr [ebp-40]
'007218a0    51                      push ecx
'007218a1    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'007218a3    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58, var_61)
'007218a9    8b16                    mov edx, dword ptr [esi]
'007218ab    83c424                  add esp, 24
'007218ae    56                      push esi
'007218af    ff9204030000            call dword ptr [edx+00000304]

' *** Reference to "__vbaObjSet"
'007218b5    8b1db4104000            mov ebx, dword ptr [004010b4]
'007218bb    50                      push eax
'007218bc    8d45bc                  lea eax, dword ptr [ebp-44]
'007218bf    50                      push eax
'007218c0    ffd3                    call ebx
Set var_58 = 
'007218c2    8b0e                    mov ecx, dword ptr [esi]
'007218c4    56                      push esi
'007218c5    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'007218cb    ff9104030000            call dword ptr [ecx+00000304]
'007218d1    50                      push eax
'007218d2    8d55c0                  lea edx, dword ptr [ebp-40]
'007218d5    52                      push edx
'007218d6    ffd3                    call ebx
Set var_5 = var_58
'007218d8    8d4de8                  lea ecx, dword ptr [ebp-18]
'007218db    8bd8                    mov ebx, eax
'007218dd    8b03                    mov eax, dword ptr [ebx]
'007218df    51                      push ecx
'007218e0    53                      push ebx
'007218e1    ff90a0000000            call dword ptr [eax+000000a0]
'007218e7    dbe2                    fnclex
'007218e9    85c0                    test ax, ax
'007218eb    7d12                    jge 7218ff
'007218ed    68a0000000              push 000000a0
'007218f2    68d00d4300              push 00430dd0
'007218f7    53                      push ebx
'007218f8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007218f9    ff1580104000            call dword ptr [00401080]
'007218ff    8b45e8                  mov eax, dword ptr [ebp-18]
'00721902    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'00721908    8b1a                    mov ebx, dword ptr [edx]
'0072190a    50                      push eax
'0072190b    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'00721910    ff1570104000            call dword ptr [00401070]
var_1473 = () & (vbCrLf)
'00721916    8bd0                    mov edx, eax
'00721918    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0072191b    ffd7                    call edi
'0072191d    50                      push eax
'0072191e    68ec694500              push 004569ec

' *** Reference to "__vbaStrCat"
'00721923    ff1570104000            call dword ptr [00401070]
var_2112 = (var_1473) & ("Mélange de races")
'00721929    8bd0                    mov edx, eax
'0072192b    8d4de0                  lea ecx, dword ptr [ebp-20]
'0072192e    ffd7                    call edi
'00721930    50                      push eax
'00721931    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'00721936    ff1570104000            call dword ptr [00401070]
var_1671 = (var_2112) & (vbCrLf)
'0072193c    8bd0                    mov edx, eax
'0072193e    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00721941    ffd7                    call edi
'00721943    8b8d20ffffff            mov ecx, dword ptr [ebp+ffffff20]
'00721949    50                      push eax
'0072194a    51                      push ecx
'0072194b    ff93a4000000            call dword ptr [ebx+000000a4]
'00721951    dbe2                    fnclex
'00721953    33db                    xor ebx, ebx
var_num2 = Empty
'00721955    3bc3                    cmp eax, ebx
'00721957    7d18                    jge 721971

If (-4944 < var_58) Then
'00721959    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'0072195f    68a4000000              push 000000a4
'00721964    68d00d4300              push 00430dd0
'00721969    52                      push edx
'0072196a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072196b    ff1580104000            call dword ptr [00401080]
    
End If
'00721971    8d45dc                  lea eax, dword ptr [ebp-24]
'00721974    50                      push eax
'00721975    8d4de0                  lea ecx, dword ptr [ebp-20]
'00721978    51                      push ecx
'00721979    8d55e4                  lea edx, dword ptr [ebp-1c]
'0072197c    52                      push edx
'0072197d    8d45e8                  lea eax, dword ptr [ebp-18]
'00721980    50                      push eax
'00721981    6a04                    push 04

' *** Reference to "__vbaFreeStrList"
'00721983    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , -4936, -4940, -4944)
'00721989    8d4dbc                  lea ecx, dword ptr [ebp-44]
'0072198c    51                      push ecx
'0072198d    8d55c0                  lea edx, dword ptr [ebp-40]
'00721990    52                      push edx
'00721991    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00721993    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58)
'00721999    8b06                    mov eax, dword ptr [esi]
'0072199b    83c420                  add esp, 20
'0072199e    56                      push esi
'0072199f    ff9014030000            call dword ptr [eax+00000314]
'007219a5    50                      push eax
'007219a6    8d4dc0                  lea ecx, dword ptr [ebp-40]
'007219a9    51                      push ecx

' *** Reference to "__vbaObjSet"
'007219aa    ff15b4104000            call dword ptr [004010b4]
Set var_5 = Nothing
'007219b0    8b10                    mov edx, dword ptr [eax]
'007219b2    8d4de8                  lea ecx, dword ptr [ebp-18]
'007219b5    51                      push ecx
'007219b6    50                      push eax
'007219b7    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'007219bd    ff92a8000000            call dword ptr [edx+000000a8]
'007219c3    dbe2                    fnclex
'007219c5    3bc3                    cmp eax, ebx
'007219c7    7d18                    jge 7219e1

If (var_5 < var_58) Then
'007219c9    8b9528ffffff            mov edx, dword ptr [ebp+ffffff28]
'007219cf    68a8000000              push 000000a8
'007219d4    681cb94300              push 0043b91c
'007219d9    52                      push edx
'007219da    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007219db    ff1580104000            call dword ptr [00401080]
    
End If
'007219e1    8b06                    mov eax, dword ptr [esi]
'007219e3    56                      push esi
'007219e4    ff9018030000            call dword ptr [eax+00000318]
'007219ea    50                      push eax
'007219eb    8d4dbc                  lea ecx, dword ptr [ebp-44]
'007219ee    51                      push ecx

' *** Reference to "__vbaObjSet"
'007219ef    ff15b4104000            call dword ptr [004010b4]
Set var_58 = Nothing
'007219f5    8b10                    mov edx, dword ptr [eax]
'007219f7    8d4de0                  lea ecx, dword ptr [ebp-20]
'007219fa    51                      push ecx
'007219fb    50                      push eax
'007219fc    898520ffffff            mov dword ptr [ebp+ffffff20], eax
'00721a02    ff92a0000000            call dword ptr [edx+000000a0]
'00721a08    dbe2                    fnclex
'00721a0a    3bc3                    cmp eax, ebx
'00721a0c    7d18                    jge 721a26

If (var_58 < var_58) Then
'00721a0e    8b9520ffffff            mov edx, dword ptr [ebp+ffffff20]
'00721a14    68a0000000              push 000000a0
'00721a19    68d00d4300              push 00430dd0
'00721a1e    52                      push edx
'00721a1f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00721a20    ff1580104000            call dword ptr [00401080]
    
End If
'00721a26    8b06                    mov eax, dword ptr [esi]
'00721a28    56                      push esi
'00721a29    ff9004030000            call dword ptr [eax+00000304]
'00721a2f    50                      push eax
'00721a30    8d4db4                  lea ecx, dword ptr [ebp-4c]
'00721a33    51                      push ecx

' *** Reference to "__vbaObjSet"
'00721a34    ff15b4104000            call dword ptr [004010b4]
Set var_62 = Nothing
'00721a3a    8b16                    mov edx, dword ptr [esi]
'00721a3c    56                      push esi
'00721a3d    89850cffffff            mov dword ptr [ebp+ffffff0c], eax
'00721a43    ff9204030000            call dword ptr [edx+00000304]
'00721a49    50                      push eax
'00721a4a    8d45b8                  lea eax, dword ptr [ebp-48]
'00721a4d    50                      push eax

' *** Reference to "__vbaObjSet"
'00721a4e    ff15b4104000            call dword ptr [004010b4]
Set var_61 = Nothing
'00721a54    8b08                    mov ecx, dword ptr [eax]
'00721a56    8d55d8                  lea edx, dword ptr [ebp-28]
'00721a59    52                      push edx
'00721a5a    50                      push eax
'00721a5b    898518ffffff            mov dword ptr [ebp+ffffff18], eax
'00721a61    ff91a0000000            call dword ptr [ecx+000000a0]
'00721a67    dbe2                    fnclex
'00721a69    3bc3                    cmp eax, ebx
'00721a6b    7d18                    jge 721a85

If (var_61 < var_58) Then
'00721a6d    8b8d18ffffff            mov ecx, dword ptr [ebp+ffffff18]
'00721a73    68a0000000              push 000000a0
'00721a78    68d00d4300              push 00430dd0
'00721a7d    51                      push ecx
'00721a7e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00721a7f    ff1580104000            call dword ptr [00401080]
    
End If
'00721a85    8b55e0                  mov edx, dword ptr [ebp-20]
'00721a88    52                      push edx

' *** Reference to "__vbaI4Str"
'00721a89    ff1550124000            call dword ptr [00401250]
'00721a8f    8b55e8                  mov edx, dword ptr [ebp-18]
'00721a92    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00721a95    898530ffffff            mov dword ptr [ebp+ffffff30], eax
'00721a9b    895de8                  mov dword ptr [ebp-18], ebx
'00721a9e    ffd7                    call edi
'00721aa0    8b06                    mov eax, dword ptr [esi]
'00721aa2    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00721aa5    51                      push ecx
'00721aa6    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'00721aac    52                      push edx
'00721aad    8d4de4                  lea ecx, dword ptr [ebp-1c]
'00721ab0    51                      push ecx
'00721ab1    56                      push esi
'00721ab2    ff9020070000            call dword ptr [eax+00000720]
'00721ab8    3bc3                    cmp eax, ebx
'00721aba    7d12                    jge 721ace

If (frmGenerateurVille < var_58) Then
'00721abc    6820070000              push 00000720
'00721ac1    68b0154300              push 004315b0
'00721ac6    56                      push esi
'00721ac7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00721ac8    ff1580104000            call dword ptr [00401080]
    
End If
'00721ace    8b45d8                  mov eax, dword ptr [ebp-28]
'00721ad1    8b4ddc                  mov ecx, dword ptr [ebp-24]
'00721ad4    8b950cffffff            mov edx, dword ptr [ebp+ffffff0c]
'00721ada    8b32                    mov esi, dword ptr [edx]
'00721adc    50                      push eax
'00721add    51                      push ecx

' *** Reference to "__vbaStrCat"
'00721ade    ff1570104000            call dword ptr [00401070]
var_1663 = (var_2342) & (var_1671)
'00721ae4    8bd0                    mov edx, eax
'00721ae6    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00721ae9    ffd7                    call edi
'00721aeb    50                      push eax
'00721aec    6870084300              push 00430870

' *** Reference to "__vbaStrCat"
'00721af1    ff1570104000            call dword ptr [00401070]
var_2109 = (var_1663) & (vbCrLf)
'00721af7    8bd0                    mov edx, eax
'00721af9    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00721afc    ffd7                    call edi
'00721afe    8bd6                    mov edx, esi
'00721b00    8bb50cffffff            mov esi, dword ptr [ebp+ffffff0c]
'00721b06    50                      push eax
'00721b07    56                      push esi
'00721b08    ff92a4000000            call dword ptr [edx+000000a4]
'00721b0e    dbe2                    fnclex
'00721b10    3bc3                    cmp eax, ebx
'00721b12    7d12                    jge 721b26

If (-4956 < var_58) Then
'00721b14    68a4000000              push 000000a4
'00721b19    68d00d4300              push 00430dd0
'00721b1e    56                      push esi
'00721b1f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00721b20    ff1580104000            call dword ptr [00401080]
    
End If
'00721b26    8d45d0                  lea eax, dword ptr [ebp-30]
'00721b29    50                      push eax
'00721b2a    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'00721b2d    51                      push ecx
'00721b2e    8d55dc                  lea edx, dword ptr [ebp-24]
'00721b31    52                      push edx
'00721b32    8d45d8                  lea eax, dword ptr [ebp-28]
'00721b35    50                      push eax
'00721b36    8d4de0                  lea ecx, dword ptr [ebp-20]
'00721b39    51                      push ecx
'00721b3a    8d55e4                  lea edx, dword ptr [ebp-1c]
'00721b3d    52                      push edx
'00721b3e    6a06                    push 06

' *** Reference to "__vbaFreeStrList"
'00721b40    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 4, -4940, -4912, -4944, -4952, -4956)
'00721b46    8d45b4                  lea eax, dword ptr [ebp-4c]
'00721b49    50                      push eax
'00721b4a    8d4db8                  lea ecx, dword ptr [ebp-48]
'00721b4d    51                      push ecx
'00721b4e    8d55bc                  lea edx, dword ptr [ebp-44]
'00721b51    52                      push edx
'00721b52    8d45c0                  lea eax, dword ptr [ebp-40]
'00721b55    50                      push eax
'00721b56    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'00721b58    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58, var_61, var_62)
'00721b5e    83c430                  add esp, 30
'00721b61    895dfc                  mov dword ptr [ebp-04], ebx
'00721b64    9b                      fwait
'00721b65    68d71b7200              push 00721bd7
'00721b6a    eb6a                    jmp 721bd6
'00721b6c    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'00721b6f    51                      push ecx
'00721b70    8d55c8                  lea edx, dword ptr [ebp-38]
'00721b73    52                      push edx
'00721b74    8d45cc                  lea eax, dword ptr [ebp-34]
'00721b77    50                      push eax
'00721b78    8d4dd0                  lea ecx, dword ptr [ebp-30]
'00721b7b    51                      push ecx
'00721b7c    8d55d4                  lea edx, dword ptr [ebp-2c]
'00721b7f    52                      push edx
'00721b80    8d45d8                  lea eax, dword ptr [ebp-28]
'00721b83    50                      push eax
'00721b84    8d4ddc                  lea ecx, dword ptr [ebp-24]
'00721b87    51                      push ecx
'00721b88    8d55e0                  lea edx, dword ptr [ebp-20]
'00721b8b    52                      push edx
'00721b8c    8d45e4                  lea eax, dword ptr [ebp-1c]
'00721b8f    50                      push eax
'00721b90    8d4de8                  lea ecx, dword ptr [ebp-18]
'00721b93    51                      push ecx
'00721b94    6a0a                    push 0a

' *** Reference to "__vbaFreeStrList"
'00721b96    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_58, 4, -4940, -4944, -4912, -4952, -4956, -4888)
'00721b9c    8d55b4                  lea edx, dword ptr [ebp-4c]
'00721b9f    52                      push edx
'00721ba0    8d45b8                  lea eax, dword ptr [ebp-48]
'00721ba3    50                      push eax
'00721ba4    8d4dbc                  lea ecx, dword ptr [ebp-44]
'00721ba7    51                      push ecx
'00721ba8    8d55c0                  lea edx, dword ptr [ebp-40]
'00721bab    52                      push edx
'00721bac    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'00721bae    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_5, var_58, var_61, var_62)
'00721bb4    83c440                  add esp, 40
'00721bb7    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'00721bbd    50                      push eax
'00721bbe    8d4d84                  lea ecx, dword ptr [ebp-7c]
'00721bc1    51                      push ecx
'00721bc2    8d5594                  lea edx, dword ptr [ebp-6c]
'00721bc5    52                      push edx
'00721bc6    8d45a4                  lea eax, dword ptr [ebp-5c]
'00721bc9    50                      push eax
'00721bca    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'00721bcc    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_83, var_121, var_119, var_91)
'00721bd2    83c414                  add esp, 14
'00721bd5    c3                      ret
'00721bd6    c3                      ret
'00721bd7    8b4508                  mov eax, dword ptr [ebp+08]
'00721bda    8b08                    mov ecx, dword ptr [eax]
'00721bdc    50                      push eax
'00721bdd    ff5108                  call dword ptr [ecx+08]
'00721be0    8b45fc                  mov eax, dword ptr [ebp-04]
'00721be3    8b4dec                  mov ecx, dword ptr [ebp-14]
'00721be6    5f                      pop edi
'00721be7    5e                      pop esi
    'Reference to '__except_list'
'00721be8    64890d00000000          mov dword ptr fs:[00000000], ecx
'00721bef    5b                      pop ebx
'00721bf0    8be5                    mov esp, ebp
'00721bf2    5d                      pop ebp
'00721bf3    c20400                  ret 0004


End Sub


'Event for Form
Private Sub Form_Load()
'0071f580    55                      push ebp
'0071f581    8bec                    mov ebp, esp
'0071f583    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0071f586    6866724000              push 00407266
'0071f58b    64a100000000            mov ax, word ptr fs:[00000000]
'0071f591    50                      push eax
    'Reference to '__except_list'
'0071f592    64892500000000          mov dword ptr fs:[00000000], esp
'0071f599    83ec14                  sub esp, 14
'0071f59c    53                      push ebx
'0071f59d    56                      push esi
'0071f59e    57                      push edi
'0071f59f    8965f4                  mov dword ptr [ebp-0c], esp
'0071f5a2    c745f8b86f4000          mov dword ptr [ebp-08], 00406fb8
'0071f5a9    8b7508                  mov esi, dword ptr [ebp+08]
'0071f5ac    8bc6                    mov eax, esi
'0071f5ae    83e001                  and eax, 01
'0071f5b1    8945fc                  mov dword ptr [ebp-04], eax
'0071f5b4    83e6fe                  and esi, fffffffe
'0071f5b7    8b0e                    mov ecx, dword ptr [esi]
'0071f5b9    56                      push esi
'0071f5ba    897508                  mov dword ptr [ebp+08], esi
'0071f5bd    ff5104                  call dword ptr [ecx+04]
'0071f5c0    8b16                    mov edx, dword ptr [esi]
'0071f5c2    33ff                    xor edi, edi
'0071f5c4    56                      push esi
'0071f5c5    897de8                  mov dword ptr [ebp-18], edi
'0071f5c8    ff9214030000            call dword ptr [edx+00000314]
'0071f5ce    50                      push eax
'0071f5cf    8d45e8                  lea eax, dword ptr [ebp-18]
'0071f5d2    50                      push eax

' *** Reference to "__vbaObjSet"
'0071f5d3    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Me
'0071f5d9    8bf0                    mov esi, eax
'0071f5db    8b0e                    mov ecx, dword ptr [esi]
'0071f5dd    57                      push edi
'0071f5de    56                      push esi
'0071f5df    ff91f4000000            call dword ptr [ecx+000000f4]
'0071f5e5    dbe2                    fnclex
'0071f5e7    3bc7                    cmp eax, edi
'0071f5e9    7d12                    jge 71f5fd

If (var_41 < 0) Then
'0071f5eb    68f4000000              push 000000f4
'0071f5f0    681cb94300              push 0043b91c
'0071f5f5    56                      push esi
'0071f5f6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0071f5f7    ff1580104000            call dword ptr [00401080]
    
End If
'0071f5fd    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0071f600    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0071f606    897dfc                  mov dword ptr [ebp-04], edi
'0071f609    681bf67100              push 0071f61b
'0071f60e    eb0a                    jmp 71f61a
'0071f610    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0071f613    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0071f619    c3                      ret
'0071f61a    c3                      ret
'0071f61b    8b4508                  mov eax, dword ptr [ebp+08]
'0071f61e    8b10                    mov edx, dword ptr [eax]
'0071f620    50                      push eax
'0071f621    ff5208                  call dword ptr [edx+08]
'0071f624    8b45fc                  mov eax, dword ptr [ebp-04]
'0071f627    8b4dec                  mov ecx, dword ptr [ebp-14]
'0071f62a    5f                      pop edi
'0071f62b    5e                      pop esi
    'Reference to '__except_list'
'0071f62c    64890d00000000          mov dword ptr fs:[00000000], ecx
'0071f633    5b                      pop ebx
'0071f634    8be5                    mov esp, ebp
'0071f636    5d                      pop ebp
'0071f637    c20400                  ret 0004


End Sub


Private Sub Form_Resize()
'00724490    55                      push ebp
'00724491    8bec                    mov ebp, esp
'00724493    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'00724496    6866724000              push 00407266
'0072449b    64a100000000            mov ax, word ptr fs:[00000000]
'007244a1    50                      push eax
    'Reference to '__except_list'
'007244a2    64892500000000          mov dword ptr fs:[00000000], esp
'007244a9    83ec70                  sub esp, 70
'007244ac    53                      push ebx
'007244ad    56                      push esi
'007244ae    57                      push edi
'007244af    8965f4                  mov dword ptr [ebp-0c], esp
'007244b2    c745f858714000          mov dword ptr [ebp-08], 00407158
'007244b9    8b7508                  mov esi, dword ptr [ebp+08]
'007244bc    8bc6                    mov eax, esi
'007244be    83e001                  and eax, 01
'007244c1    8945fc                  mov dword ptr [ebp-04], eax
'007244c4    83e6fe                  and esi, fffffffe
'007244c7    8b0e                    mov ecx, dword ptr [esi]
'007244c9    56                      push esi
'007244ca    897508                  mov dword ptr [ebp+08], esi
'007244cd    ff5104                  call dword ptr [ecx+04]
'007244d0    8b16                    mov edx, dword ptr [esi]
'007244d2    33ff                    xor edi, edi
'007244d4    56                      push esi
'007244d5    897de8                  mov dword ptr [ebp-18], edi
'007244d8    897dd8                  mov dword ptr [ebp-28], edi
'007244db    897dc8                  mov dword ptr [ebp-38], edi
'007244de    897da4                  mov dword ptr [ebp-5c], edi
'007244e1    ff9204030000            call dword ptr [edx+00000304]
'007244e7    50                      push eax
'007244e8    8d45e8                  lea eax, dword ptr [ebp-18]
'007244eb    50                      push eax

' *** Reference to "__vbaObjSet"
'007244ec    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Me
'007244f2    8bd8                    mov ebx, eax
'007244f4    393dfcb07200            cmp dword ptr [0072b0fc], edi
'007244fa    7510                    jne 72450c
'007244fc    68fcb07200              push 0072b0fc
'00724501    687cbb4000              push 0040bb7c

' *** Reference to "__vbaNew2"
'00724506    ff152c124000            call dword ptr [0040122c]
Dim var_34 As New frmGenerateurVille
'0072450c    8b3dfcb07200            mov edi, dword ptr [0072b0fc]
'00724512    8b0f                    mov ecx, dword ptr [edi]
'00724514    8d55a4                  lea edx, dword ptr [ebp-5c]
'00724517    52                      push edx
'00724518    57                      push edi
'00724519    ff9188000000            call dword ptr [ecx+00000088]
'0072451f    dbe2                    fnclex
'00724521    85c0                    test ax, ax
'00724523    7d12                    jge 724537
'00724525    6888000000              push 00000088
'0072452a    6880154300              push 00431580
'0072452f    57                      push edi
'00724530    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00724531    ff1580104000            call dword ptr [00401080]
'00724537    d945a4                  fld dword ptr [ebp-5c]
'0072453a    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0072453d    d82550714000            fsub dword ptr [00407150]
'00724543    c745c804000000          mov dword ptr [ebp-38], 00000004
'0072454a    c745e001000000          mov dword ptr [ebp-20], 00000001
'00724551    c745d802000000          mov dword ptr [ebp-28], 00000002
'00724558    d95dd0                  fstp dword ptr [ebp-30]
'var_4 = (00)
'0072455b    dfe0                    fnstsw ax
'0072455d    a80d                    test al, 0d
'0072455f    0f8580010000            jne 7246e5
'00724565    8b3b                    mov edi, dword ptr [ebx]
'00724567    8d45c8                  lea eax, dword ptr [ebp-38]
'0072456a    50                      push eax
'0072456b    51                      push ecx
'0072456c    e87fe1dcff              call 4f26f0
Call sub_4F26F0()
'00724571    0fbfd0                  movsx edx, ax
'00724574    895588                  mov dword ptr [ebp-78], edx
'00724577    db4588                  fild dword ptr [ebp-78]
'0072457a    d95d84                  fstp dword ptr [ebp-7c]
'var_119 = (00)
'0072457d    8b4584                  mov eax, dword ptr [ebp-7c]
'00724580    50                      push eax
'00724581    53                      push ebx
'00724582    ff9784000000            call dword ptr [edi+00000084]
'00724588    dbe2                    fnclex
'0072458a    85c0                    test ax, ax
'0072458c    7d12                    jge 7245a0
'0072458e    6884000000              push 00000084
'00724593    68d00d4300              push 00430dd0
'00724598    53                      push ebx
'00724599    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072459a    ff1580104000            call dword ptr [00401080]

' *** Reference to "__vbaFreeObj"
'007245a0    8b1d08134000            mov ebx, dword ptr [00401308]
'007245a6    8d4de8                  lea ecx, dword ptr [ebp-18]
'007245a9    ffd3                    call ebx
'#Cleanup(var_41)
'007245ab    8d4dc8                  lea ecx, dword ptr [ebp-38]
'007245ae    51                      push ecx
'007245af    8d55d8                  lea edx, dword ptr [ebp-28]
'007245b2    52                      push edx
'007245b3    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'007245b5    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_45, var_46)
'007245bb    8b06                    mov eax, dword ptr [esi]
'007245bd    83c40c                  add esp, 0c
'007245c0    56                      push esi
'007245c1    ff9004030000            call dword ptr [eax+00000304]
'007245c7    50                      push eax
'007245c8    8d4de8                  lea ecx, dword ptr [ebp-18]
'007245cb    51                      push ecx

' *** Reference to "__vbaObjSet"
'007245cc    ff15b4104000            call dword ptr [004010b4]
Set var_41 = Nothing
'007245d2    8bf0                    mov esi, eax
'007245d4    a1fcb07200              mov ax, word ptr [0072b0fc]
'007245d9    85c0                    test ax, ax
'007245db    7510                    jne 7245ed
'007245dd    68fcb07200              push 0072b0fc
'007245e2    687cbb4000              push 0040bb7c

' *** Reference to "__vbaNew2"
'007245e7    ff152c124000            call dword ptr [0040122c]
Set var_34 = New frmGenerateurVille
'007245ed    8b3dfcb07200            mov edi, dword ptr [0072b0fc]
'007245f3    8b17                    mov edx, dword ptr [edi]
'007245f5    8d45a4                  lea eax, dword ptr [ebp-5c]
'007245f8    50                      push eax
'007245f9    57                      push edi
'007245fa    ff9280000000            call dword ptr [edx+00000080]
'00724600    dbe2                    fnclex
'00724602    85c0                    test ax, ax
'00724604    7d12                    jge 724618
'00724606    6880000000              push 00000080
'0072460b    6880154300              push 00431580
'00724610    57                      push edi
'00724611    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00724612    ff1580104000            call dword ptr [00401080]
'00724618    d945a4                  fld dword ptr [ebp-5c]
'0072461b    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0072461e    d82540554000            fsub dword ptr [00405540]
'00724624    51                      push ecx
'00724625    8d55d8                  lea edx, dword ptr [ebp-28]
'00724628    c745c804000000          mov dword ptr [ebp-38], 00000004
'0072462f    d95dd0                  fstp dword ptr [ebp-30]
'var_4 = (00)
'00724632    dfe0                    fnstsw ax
'00724634    a80d                    test al, 0d
'00724636    0f85a9000000            jne 7246e5
'0072463c    c745e001000000          mov dword ptr [ebp-20], 00000001
'00724643    c745d802000000          mov dword ptr [ebp-28], 00000002
'0072464a    8b3e                    mov edi, dword ptr [esi]
'0072464c    52                      push edx
'0072464d    e89ee0dcff              call 4f26f0
Call sub_4F26F0()
'00724652    0fbfc0                  movsx eax, ax
'00724655    894580                  mov dword ptr [ebp-80], eax
'00724658    db4580                  fild dword ptr [ebp-80]
'0072465b    d99d7cffffff            fstp dword ptr [ebp+ffffff7c]
'var_59 = (00)
'00724661    8b8d7cffffff            mov ecx, dword ptr [ebp+ffffff7c]
'00724667    51                      push ecx
'00724668    56                      push esi
'00724669    ff577c                  call dword ptr [edi+7c]
'0072466c    dbe2                    fnclex
'0072466e    85c0                    test ax, ax
'00724670    7d0f                    jge 724681
'00724672    6a7c                    push 7c
'00724674    68d00d4300              push 00430dd0
'00724679    56                      push esi
'0072467a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0072467b    ff1580104000            call dword ptr [00401080]
'00724681    8d4de8                  lea ecx, dword ptr [ebp-18]
'00724684    ffd3                    call ebx
'#Cleanup(var_41)
'00724686    8d55c8                  lea edx, dword ptr [ebp-38]
'00724689    52                      push edx
'0072468a    8d45d8                  lea eax, dword ptr [ebp-28]
'0072468d    50                      push eax
'0072468e    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'00724690    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_45, var_46)
'00724696    83c40c                  add esp, 0c
'00724699    c745fc00000000          mov dword ptr [ebp-04], 00000000
'007246a0    9b                      fwait
'007246a1    68c6467200              push 007246c6
'007246a6    eb1d                    jmp 7246c5
'007246a8    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'007246ab    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'007246b1    8d4dc8                  lea ecx, dword ptr [ebp-38]
'007246b4    51                      push ecx
'007246b5    8d55d8                  lea edx, dword ptr [ebp-28]
'007246b8    52                      push edx
'007246b9    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'007246bb    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_45, var_46)
'007246c1    83c40c                  add esp, 0c
'007246c4    c3                      ret
'007246c5    c3                      ret
'007246c6    8b4508                  mov eax, dword ptr [ebp+08]
'007246c9    8b08                    mov ecx, dword ptr [eax]
'007246cb    50                      push eax
'007246cc    ff5108                  call dword ptr [ecx+08]
'007246cf    8b45fc                  mov eax, dword ptr [ebp-04]
'007246d2    8b4dec                  mov ecx, dword ptr [ebp-14]
'007246d5    5f                      pop edi
'007246d6    5e                      pop esi
    'Reference to '__except_list'
'007246d7    64890d00000000          mov dword ptr fs:[00000000], ecx
'007246de    5b                      pop ebx
'007246df    8be5                    mov esp, ebp
'007246e1    5d                      pop ebp
'007246e2    c20400                  ret 0004


End Sub


