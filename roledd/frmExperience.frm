VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "C:\WINDOWS\SysWow64\Vsflex7d.ocx"

Begin VB.Form frmExperience
    Caption = "Expérience des personnages"
    ScaleMode = 1
    AutoRedraw = 0              'False
    FontTransparent = -1              'True
    LinkTopic = "Form1"
    MDIChild = -1              'True
    ClientLeft   = 60
    ClientTop    = 450
    ClientWidth  = 11580
    ClientHeight = 7485
    BeginProperty Font
        Name          = "Times New Roman"
        Size          = 9,75
        Charset       = 0
        Weight        = 700
        Underline     = 0              'False
        Italic        = 0              'False
        Strikethrough = 0              'False
    EndProperty
    Begin VB.TextBox fldnIdee
        Left   = 9000
        Top    = 3720
        Width  = 615
        Height = 375
        Text = "0"
        TabIndex = 7
        ToolTipText = "Saisir un pourcentage pour les XP idée."
    End
    Begin VB.CommandButton btnAdd
        Caption = "Ajouter"
        Left   = 9720
        Top    = 2760
        Width  = 855
        Height = 375
        TabIndex = 5
    End
    Begin VB.TextBox fldnFP
        Left   = 9000
        Top    = 2760
        Width  = 615
        Height = 375
        TabIndex = 4
    End
    Begin VB.CommandButton btnFermer
        Caption = "Fermer"
        Left   = 10080
        Top    = 6960
        Width  = 1335
        Height = 375
        TabIndex = 3
    End
    Begin VB.TextBox fldstrExperience
        Left   = 0
        Top    = 4800
        Width  = 8565
        Height = 2595
        TabIndex = 2
        MultiLine = -1              'True
        ScrollBars = 2
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 12
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
    Begin VB.CommandButton btnEnregistrer
        Caption = "Enregistrer"
        Left   = 8640
        Top    = 6960
        Width  = 1335
        Height = 375
        TabIndex = 1
    End
    Begin VSFlex7DAOCtl.VSFlexGrid vspersonnage
        Left   = 0
        Top    = 0
        Width  = 8580
        Height = 4740
        TabIndex = 0
        OleObjectBlob = frmExperience.frx:0000
    End
    Begin VB.Label fldstrTexteIdee
        Caption = "Total XP idée"
        Index = 1
        Left   = 9960
        Top    = 3240
        Width  = 1095
        Height = 255
        TabIndex = 10
    End
    Begin VB.Label fldstrTexteIdee
        Caption = "% d'XP idée"
        Index = 0
        Left   = 8760
        Top    = 3240
        Width  = 1215
        Height = 255
        TabIndex = 9
    End
    Begin VB.Label fldnTotalIdee
        Left   = 10080
        Top    = 3720
        Width  = 975
        Height = 375
        TabIndex = 8
    End
    Begin VB.Label Label1
        Caption = "Sélectionner les personnages (ctrl+click) puis saisir le FP de la rencontre et enfin cliquer sur ajouter. Enregistrer pour valider. Part vaut 1 pour un PJ et 0.5 pour un PNJ."
        Left   = 8760
        Top    = 240
        Width  = 2655
        Height = 2175
        TabIndex = 6
        BeginProperty Font
            Name          = "Times New Roman"
            Size          = 12
            Charset       = 0
            Weight        = 400
            Underline     = 0              'False
            Italic        = 0              'False
            Strikethrough = 0              'False
        EndProperty
    End
End
Public Function RempPerso(arg_0 As Unknow, arg_1 As Unknow, arg_2 As Unknow, arg_3 As Unknow, arg_4 As Unknow, arg_5 As Unknow, arg_6 As Unknow, arg_7 As Unknow, arg_8 As Unknow, arg_9 As Unknow, arg_A As Unknow, arg_B As Unknow, arg_C As Unknow, arg_D As Unknow, arg_E As Unknow, arg_F As Unknow, arg_10 As Unknow, arg_11 As Unknow, arg_12 As Unknow, arg_13 As Unknow, arg_14 As Unknow, arg_15 As Unknow, arg_16 As Unknow, arg_17 As Unknow, arg_18 As Unknow, arg_19 As Unknow, arg_1A As Unknow, arg_1B As Unknow, arg_1C As Unknow, arg_1D As Unknow, arg_1E As Unknow, arg_1F As Unknow, arg_20 As Unknow, arg_21 As Unknow, arg_22 As Unknow, arg_23 As Unknow, arg_24 As Unknow, arg_25 As Unknow, arg_26 As Unknow, arg_27 As Unknow, arg_28 As Unknow, arg_29 As Unknow, arg_2A As Unknow, arg_2B As Unknow, arg_2C As Unknow, arg_2D As Unknow, arg_2E As Unknow, arg_2F As Unknow, arg_30 As Unknow, arg_31 As Unknow, arg_32 As Unknow, arg_33 As Unknow, arg_34 As Unknow, arg_35 As Unknow, arg_36 As Unknow, arg_37 As Unknow, arg_38 As Unknow, arg_39 As Unknow, arg_3A As Unknow, arg_3B As Unknow, arg_3C As Unknow)
'00708270    55                      push ebp
'00708271    8bec                    mov ebp, esp
'00708273    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'00708276    6866724000              push 00407266
'0070827b    64a100000000            mov ax, word ptr fs:[00000000]
'00708281    50                      push eax
    'Reference to '__except_list'
'00708282    64892500000000          mov dword ptr fs:[00000000], esp
'00708289    81ecf4020000            sub esp, 000002f4
'0070828f    53                      push ebx
'00708290    56                      push esi
'00708291    57                      push edi
'00708292    8965ec                  mov dword ptr [ebp-14], esp
'00708295    c745f0006c4000          mov dword ptr [ebp-10], 00406c00
'0070829c    33db                    xor ebx, ebx
'0070829e    895df4                  mov dword ptr [ebp-0c], ebx
'007082a1    895df8                  mov dword ptr [ebp-08], ebx
'007082a4    8b7d08                  mov edi, dword ptr [ebp+08]
'007082a7    8b07                    mov eax, dword ptr [edi]
'007082a9    57                      push edi
'007082aa    ff5004                  call dword ptr [eax+04]
'007082ad    895dd8                  mov dword ptr [ebp-28], ebx
'007082b0    895dc4                  mov dword ptr [ebp-3c], ebx
'007082b3    895db8                  mov dword ptr [ebp-48], ebx
'007082b6    895da4                  mov dword ptr [ebp-5c], ebx
'007082b9    895d94                  mov dword ptr [ebp-6c], ebx
'007082bc    895d90                  mov dword ptr [ebp-70], ebx
'007082bf    895d8c                  mov dword ptr [ebp-74], ebx
'007082c2    895d88                  mov dword ptr [ebp-78], ebx
'007082c5    895d84                  mov dword ptr [ebp-7c], ebx
'007082c8    895d80                  mov dword ptr [ebp-80], ebx
'007082cb    899d7cffffff            mov dword ptr [ebp+ffffff7c], ebx
'007082d1    899d78ffffff            mov dword ptr [ebp+ffffff78], ebx
'007082d7    899d74ffffff            mov dword ptr [ebp+ffffff74], ebx
'007082dd    899d70ffffff            mov dword ptr [ebp+ffffff70], ebx
'007082e3    899d6cffffff            mov dword ptr [ebp+ffffff6c], ebx
'007082e9    899d5cffffff            mov dword ptr [ebp+ffffff5c], ebx
'007082ef    899d4cffffff            mov dword ptr [ebp+ffffff4c], ebx
'007082f5    899d3cffffff            mov dword ptr [ebp+ffffff3c], ebx
'007082fb    899d2cffffff            mov dword ptr [ebp+ffffff2c], ebx
'00708301    899d1cffffff            mov dword ptr [ebp+ffffff1c], ebx
'00708307    899d0cffffff            mov dword ptr [ebp+ffffff0c], ebx
'0070830d    899dfcfeffff            mov dword ptr [ebp+fffffefc], ebx
'00708313    899decfeffff            mov dword ptr [ebp+fffffeec], ebx
'00708319    899ddcfeffff            mov dword ptr [ebp+fffffedc], ebx
'0070831f    899dccfeffff            mov dword ptr [ebp+fffffecc], ebx
'00708325    899dbcfeffff            mov dword ptr [ebp+fffffebc], ebx
'0070832b    899dacfeffff            mov dword ptr [ebp+fffffeac], ebx
'00708331    899d9cfeffff            mov dword ptr [ebp+fffffe9c], ebx
'00708337    899d8cfeffff            mov dword ptr [ebp+fffffe8c], ebx
'0070833d    899d7cfeffff            mov dword ptr [ebp+fffffe7c], ebx
'00708343    899d6cfeffff            mov dword ptr [ebp+fffffe6c], ebx
'00708349    899d5cfeffff            mov dword ptr [ebp+fffffe5c], ebx
'0070834f    899d4cfeffff            mov dword ptr [ebp+fffffe4c], ebx
'00708355    899d3cfeffff            mov dword ptr [ebp+fffffe3c], ebx
'0070835b    899d2cfeffff            mov dword ptr [ebp+fffffe2c], ebx
'00708361    899d1cfeffff            mov dword ptr [ebp+fffffe1c], ebx
'00708367    899d0cfeffff            mov dword ptr [ebp+fffffe0c], ebx
'0070836d    899dfcfdffff            mov dword ptr [ebp+fffffdfc], ebx
'00708373    899decfdffff            mov dword ptr [ebp+fffffdec], ebx
'00708379    899ddcfdffff            mov dword ptr [ebp+fffffddc], ebx
'0070837f    899dccfdffff            mov dword ptr [ebp+fffffdcc], ebx
'00708385    899dbcfdffff            mov dword ptr [ebp+fffffdbc], ebx
'0070838b    899dacfdffff            mov dword ptr [ebp+fffffdac], ebx
'00708391    899d9cfdffff            mov dword ptr [ebp+fffffd9c], ebx
'00708397    899d98fdffff            mov dword ptr [ebp+fffffd98], ebx
'0070839d    899d94fdffff            mov dword ptr [ebp+fffffd94], ebx
'007083a3    66391d28b07200          cmp word ptr [0072b028], bx
'007083aa    7508                    jne 7083b4
'007083ac    6a01                    push 01

' *** Reference to "__vbaOnError"
'007083ae    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'007083b4    b801000000              mov eax, 00000001
'007083b9    898594feffff            mov dword ptr [ebp+fffffe94], eax
'007083bf    b903000000              mov ecx, 00000003
'007083c4    898d8cfeffff            mov dword ptr [ebp+fffffe8c], ecx
'007083ca    83ec10                  sub esp, 10
'007083cd    8bd4                    mov edx, esp
'007083cf    890a                    mov dword ptr [edx], ecx
'007083d1    8b8d90feffff            mov ecx, dword ptr [ebp+fffffe90]
'007083d7    894a04                  mov dword ptr [edx+04], ecx
'007083da    894208                  mov dword ptr [edx+08], eax
'007083dd    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'007083e3    89420c                  mov dword ptr [edx+0c], eax
'007083e6    6a07                    push 07
'007083e8    8b0f                    mov ecx, dword ptr [edi]
'007083ea    57                      push edi
'007083eb    ff9120030000            call dword ptr [ecx+00000320]
'007083f1    50                      push eax
'007083f2    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'007083f8    52                      push edx

' *** Reference to "__vbaObjSet"
'007083f9    8b35b4104000            mov esi, dword ptr [004010b4]
'007083ff    ffd6                    call esi
Set var_87 = Nothing
'00708401    50                      push eax

' *** Reference to "__vbaLateIdSt"
'00708402    ff15ec124000            call dword ptr [004012ec]
var_87.[UNMANAGED] = 1
'00708408    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeObj"
'0070840e    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_87)
'00708414    b904000280              mov ecx, 80020004
'00708419    898d74feffff            mov dword ptr [ebp+fffffe74], ecx
'0070841f    b80a000000              mov eax, 0000000a
'00708424    89856cfeffff            mov dword ptr [ebp+fffffe6c], eax
'0070842a    898d84feffff            mov dword ptr [ebp+fffffe84], ecx
'00708430    89857cfeffff            mov dword ptr [ebp+fffffe7c], eax
'00708436    c78594feffff02000000    mov dword ptr [ebp+fffffe94], 00000002
'00708440    c7858cfeffff03000000    mov dword ptr [ebp+fffffe8c], 00000003
'0070844a    8b0d48b07200            mov ecx, dword ptr [0072b048]
'00708450    8b11                    mov edx, dword ptr [ecx]
'00708452    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'00708458    51                      push ecx
'00708459    83ec10                  sub esp, 10
'0070845c    8bcc                    mov ecx, esp
'0070845e    8901                    mov dword ptr [ecx], eax
'00708460    8b8570feffff            mov eax, dword ptr [ebp+fffffe70]
'00708466    894104                  mov dword ptr [ecx+04], eax
'00708469    8b8574feffff            mov eax, dword ptr [ebp+fffffe74]
'0070846f    894108                  mov dword ptr [ecx+08], eax
'00708472    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'00708478    89410c                  mov dword ptr [ecx+0c], eax
'0070847b    83ec10                  sub esp, 10
'0070847e    8bcc                    mov ecx, esp
'00708480    8b857cfeffff            mov eax, dword ptr [ebp+fffffe7c]
'00708486    8901                    mov dword ptr [ecx], eax
'00708488    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'0070848e    894104                  mov dword ptr [ecx+04], eax
'00708491    8b8584feffff            mov eax, dword ptr [ebp+fffffe84]
'00708497    894108                  mov dword ptr [ecx+08], eax
'0070849a    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'007084a0    89410c                  mov dword ptr [ecx+0c], eax
'007084a3    83ec10                  sub esp, 10
'007084a6    8bcc                    mov ecx, esp
'007084a8    8b858cfeffff            mov eax, dword ptr [ebp+fffffe8c]
'007084ae    8901                    mov dword ptr [ecx], eax
'007084b0    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'007084b6    894104                  mov dword ptr [ecx+04], eax
'007084b9    8b8594feffff            mov eax, dword ptr [ebp+fffffe94]
'007084bf    894108                  mov dword ptr [ecx+08], eax
'007084c2    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'007084c8    89410c                  mov dword ptr [ecx+0c], eax
'007084cb    6848614500              push 00456148
'007084d0    8b0d48b07200            mov ecx, dword ptr [0072b048]
'007084d6    51                      push ecx
'007084d7    ff92bc000000            call dword ptr [edx+000000bc]
'007084dd    dbe2                    fnclex
'007084df    3bc3                    cmp eax, ebx
'007084e1    7d18                    jge 7084fb

If (0 < 0) Then
'007084e3    68bc000000              push 000000bc
'007084e8    68301f4300              push 00431f30
'007084ed    8b1548b07200            mov edx, dword ptr [0072b048]
'007084f3    52                      push edx
'007084f4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007084f5    ff1580104000            call dword ptr [00401080]
    
End If
'007084fb    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'00708501    899d78ffffff            mov dword ptr [ebp+ffffff78], ebx
'00708507    50                      push eax
'00708508    8d45a4                  lea eax, dword ptr [ebp-5c]
'0070850b    50                      push eax
'0070850c    ffd6                    call esi
Set var_83 = Nothing
'0070850e    b904000280              mov ecx, 80020004
'00708513    898d74feffff            mov dword ptr [ebp+fffffe74], ecx
'00708519    b80a000000              mov eax, 0000000a
'0070851e    89856cfeffff            mov dword ptr [ebp+fffffe6c], eax
'00708524    898d84feffff            mov dword ptr [ebp+fffffe84], ecx
'0070852a    89857cfeffff            mov dword ptr [ebp+fffffe7c], eax
'00708530    c78594feffff01000000    mov dword ptr [ebp+fffffe94], 00000001
'0070853a    c7858cfeffff03000000    mov dword ptr [ebp+fffffe8c], 00000003
'00708544    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'0070854a    8b11                    mov edx, dword ptr [ecx]
'0070854c    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'00708552    51                      push ecx
'00708553    83ec10                  sub esp, 10
'00708556    8bcc                    mov ecx, esp
'00708558    8901                    mov dword ptr [ecx], eax
'0070855a    8b8570feffff            mov eax, dword ptr [ebp+fffffe70]
'00708560    894104                  mov dword ptr [ecx+04], eax
'00708563    8b8574feffff            mov eax, dword ptr [ebp+fffffe74]
'00708569    894108                  mov dword ptr [ecx+08], eax
'0070856c    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'00708572    89410c                  mov dword ptr [ecx+0c], eax
'00708575    83ec10                  sub esp, 10
'00708578    8bcc                    mov ecx, esp
'0070857a    8b857cfeffff            mov eax, dword ptr [ebp+fffffe7c]
'00708580    8901                    mov dword ptr [ecx], eax
'00708582    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'00708588    894104                  mov dword ptr [ecx+04], eax
'0070858b    8b8584feffff            mov eax, dword ptr [ebp+fffffe84]
'00708591    894108                  mov dword ptr [ecx+08], eax
'00708594    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'0070859a    89410c                  mov dword ptr [ecx+0c], eax
'0070859d    83ec10                  sub esp, 10
'007085a0    8bcc                    mov ecx, esp
'007085a2    8b858cfeffff            mov eax, dword ptr [ebp+fffffe8c]
'007085a8    8901                    mov dword ptr [ecx], eax
'007085aa    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'007085b0    894104                  mov dword ptr [ecx+04], eax
'007085b3    8b8594feffff            mov eax, dword ptr [ebp+fffffe94]
'007085b9    894108                  mov dword ptr [ecx+08], eax
'007085bc    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'007085c2    89410c                  mov dword ptr [ecx+0c], eax
'007085c5    688cd04300              push 0043d08c
'007085ca    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'007085d0    51                      push ecx
'007085d1    ff92bc000000            call dword ptr [edx+000000bc]
'007085d7    dbe2                    fnclex
'007085d9    3bc3                    cmp eax, ebx
'007085db    7d18                    jge 7085f5

If (0 < 0) Then
'007085dd    68bc000000              push 000000bc
'007085e2    68301f4300              push 00431f30
'007085e7    8b154cb07200            mov edx, dword ptr [0072b04c]
'007085ed    52                      push edx
'007085ee    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007085ef    ff1580104000            call dword ptr [00401080]
    
End If
'007085f5    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'007085fb    899d78ffffff            mov dword ptr [ebp+ffffff78], ebx
'00708601    50                      push eax
'00708602    8d45c4                  lea eax, dword ptr [ebp-3c]
'00708605    50                      push eax
'00708606    ffd6                    call esi
Set var_9 = Nothing
'00708608    b904000280              mov ecx, 80020004
'0070860d    898d74feffff            mov dword ptr [ebp+fffffe74], ecx
'00708613    b80a000000              mov eax, 0000000a
'00708618    89856cfeffff            mov dword ptr [ebp+fffffe6c], eax
'0070861e    898d84feffff            mov dword ptr [ebp+fffffe84], ecx
'00708624    89857cfeffff            mov dword ptr [ebp+fffffe7c], eax
'0070862a    c78594feffff01000000    mov dword ptr [ebp+fffffe94], 00000001
'00708634    c7858cfeffff03000000    mov dword ptr [ebp+fffffe8c], 00000003
'0070863e    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'00708644    8b11                    mov edx, dword ptr [ecx]
'00708646    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0070864c    51                      push ecx
'0070864d    83ec10                  sub esp, 10
'00708650    8bcc                    mov ecx, esp
'00708652    8901                    mov dword ptr [ecx], eax
'00708654    8b8570feffff            mov eax, dword ptr [ebp+fffffe70]
'0070865a    894104                  mov dword ptr [ecx+04], eax
'0070865d    8b8574feffff            mov eax, dword ptr [ebp+fffffe74]
'00708663    894108                  mov dword ptr [ecx+08], eax
'00708666    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'0070866c    89410c                  mov dword ptr [ecx+0c], eax
'0070866f    83ec10                  sub esp, 10
'00708672    8bcc                    mov ecx, esp
'00708674    8b857cfeffff            mov eax, dword ptr [ebp+fffffe7c]
'0070867a    8901                    mov dword ptr [ecx], eax
'0070867c    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'00708682    894104                  mov dword ptr [ecx+04], eax
'00708685    8b8584feffff            mov eax, dword ptr [ebp+fffffe84]
'0070868b    894108                  mov dword ptr [ecx+08], eax
'0070868e    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'00708694    89410c                  mov dword ptr [ecx+0c], eax
'00708697    83ec10                  sub esp, 10
'0070869a    8bcc                    mov ecx, esp
'0070869c    8b858cfeffff            mov eax, dword ptr [ebp+fffffe8c]
'007086a2    8901                    mov dword ptr [ecx], eax
'007086a4    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'007086aa    894104                  mov dword ptr [ecx+04], eax
'007086ad    8b8594feffff            mov eax, dword ptr [ebp+fffffe94]
'007086b3    894108                  mov dword ptr [ecx+08], eax
'007086b6    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'007086bc    89410c                  mov dword ptr [ecx+0c], eax
'007086bf    6860624500              push 00456260
'007086c4    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'007086ca    51                      push ecx
'007086cb    ff92bc000000            call dword ptr [edx+000000bc]
'007086d1    dbe2                    fnclex
'007086d3    3bc3                    cmp eax, ebx
'007086d5    7d18                    jge 7086ef

If (0 < 0) Then
'007086d7    68bc000000              push 000000bc
'007086dc    68301f4300              push 00431f30
'007086e1    8b154cb07200            mov edx, dword ptr [0072b04c]
'007086e7    52                      push edx
'007086e8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007086e9    ff1580104000            call dword ptr [00401080]
    
End If
'007086ef    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'007086f5    899d78ffffff            mov dword ptr [ebp+ffffff78], ebx
'007086fb    50                      push eax
'007086fc    8d45b8                  lea eax, dword ptr [ebp-48]
'007086ff    50                      push eax
'00708700    ffd6                    call esi
Set var_61 = Nothing
'00708702    8b45c4                  mov eax, dword ptr [ebp-3c]
'00708705    8b08                    mov ecx, dword ptr [eax]
'00708707    6830a84300              push 0043a830
'0070870c    50                      push eax
'0070870d    ff5144                  call dword ptr [ecx+44]
'00708710    dbe2                    fnclex
'00708712    3bc3                    cmp eax, ebx
'00708714    7d12                    jge 708728

If (var_9 < 0) Then
'00708716    6a44                    push 44
'00708718    6830314300              push 00433130
'0070871d    8b55c4                  mov edx, dword ptr [ebp-3c]
'00708720    52                      push edx
'00708721    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00708722    ff1580104000            call dword ptr [00401080]
    
End If
'00708728    8b45b8                  mov eax, dword ptr [ebp-48]
'0070872b    8b08                    mov ecx, dword ptr [eax]
'0070872d    6830a84300              push 0043a830
'00708732    50                      push eax
'00708733    ff5144                  call dword ptr [ecx+44]
'00708736    dbe2                    fnclex
'00708738    3bc3                    cmp eax, ebx
'0070873a    7d12                    jge 70874e

If (var_61 < 0) Then
'0070873c    6a44                    push 44
'0070873e    6830314300              push 00433130
'00708743    8b55b8                  mov edx, dword ptr [ebp-48]
'00708746    52                      push edx
'00708747    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00708748    ff1580104000            call dword ptr [00401080]
    
End If
'0070874e    c745e001000000          mov dword ptr [ebp-20], 00000001
'00708755    8b45a4                  mov eax, dword ptr [ebp-5c]
'00708758    8b08                    mov ecx, dword ptr [eax]
'0070875a    8d9598fdffff            lea edx, dword ptr [ebp+fffffd98]
'00708760    52                      push edx
'00708761    50                      push eax
'00708762    ff5134                  call dword ptr [ecx+34]
'00708765    dbe2                    fnclex
'00708767    3bc3                    cmp eax, ebx
'00708769    7d12                    jge 70877d

If (var_83 < 0) Then
'0070876b    6a34                    push 34
'0070876d    6830314300              push 00433130
'00708772    8b4da4                  mov ecx, dword ptr [ebp-5c]
'00708775    51                      push ecx
'00708776    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00708777    ff1580104000            call dword ptr [00401080]
    
End If
'0070877d    66399d98fdffff          cmp word ptr [ebp+fffffd98], bx
'00708784    8b45a4                  mov eax, dword ptr [ebp-5c]
'00708787    0f8521310000            jne 70b8ae

Do While (0 = 0)
'0070878d    8b10                    mov edx, dword ptr [eax]
'0070878f    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'00708795    51                      push ecx
'00708796    50                      push eax
'00708797    ff92b4000000            call dword ptr [edx+000000b4]
'0070879d    dbe2                    fnclex
'0070879f    3bc3                    cmp eax, ebx
'007087a1    7d15                    jge 7087b8
    
    If (    var_83 < 0) Then
'007087a3    68b4000000              push 000000b4
'007087a8    6830314300              push 00433130
'007087ad    8b55a4                  mov edx, dword ptr [ebp-5c]
'007087b0    52                      push edx
'007087b1    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'007087b2    ff1580104000            call dword ptr [00401080]
    
End If
'007087b8    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'007087be    89858cfdffff            mov dword ptr [ebp+fffffd8c], eax
'007087c4    b9e4b14300              mov ecx, 0043b1e4
'007087c9    898d94feffff            mov dword ptr [ebp+fffffe94], ecx
'007087cf    ba08000000              mov edx, 00000008
'007087d4    89958cfeffff            mov dword ptr [ebp+fffffe8c], edx
'007087da    8b30                    mov esi, dword ptr [eax]
'007087dc    8dbd74ffffff            lea edi, dword ptr [ebp+ffffff74]
'007087e2    57                      push edi
'007087e3    83ec10                  sub esp, 10
'007087e6    8bfc                    mov edi, esp
'007087e8    8917                    mov dword ptr [edi], edx
'007087ea    8b9590feffff            mov edx, dword ptr [ebp+fffffe90]
'007087f0    895704                  mov dword ptr [edi+04], edx
'007087f3    894f08                  mov dword ptr [edi+08], ecx
'007087f6    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'007087fc    894f0c                  mov dword ptr [edi+0c], ecx
'007087ff    50                      push eax
'00708800    ff5630                  call dword ptr [esi+30]
'00708803    dbe2                    fnclex
'00708805    3bc3                    cmp eax, ebx
'00708807    7d15                    jge 70881e

If (0 < 0) Then
'00708809    6a30                    push 30
'0070880b    68d8304300              push 004330d8
'00708810    8b958cfdffff            mov edx, dword ptr [ebp+fffffd8c]
'00708816    52                      push edx
'00708817    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00708818    ff1580104000            call dword ptr [00401080]
    
End If
'0070881e    b804000280              mov eax, 80020004
'00708823    b90a000000              mov ecx, 0000000a
'00708828    8bd0                    mov edx, eax
'0070882a    8995e4fdffff            mov dword ptr [ebp+fffffde4], edx
'00708830    8bf1                    mov esi, ecx
'00708832    89b5dcfdffff            mov dword ptr [ebp+fffffddc], esi
'00708838    8bf8                    mov edi, eax
'0070883a    89bdf4fdffff            mov dword ptr [ebp+fffffdf4], edi
'00708840    898decfdffff            mov dword ptr [ebp+fffffdec], ecx
'00708846    89bd04feffff            mov dword ptr [ebp+fffffe04], edi
'0070884c    898dfcfdffff            mov dword ptr [ebp+fffffdfc], ecx
'00708852    89bd14feffff            mov dword ptr [ebp+fffffe14], edi
'00708858    898d0cfeffff            mov dword ptr [ebp+fffffe0c], ecx
'0070885e    89bd24feffff            mov dword ptr [ebp+fffffe24], edi
'00708864    898d1cfeffff            mov dword ptr [ebp+fffffe1c], ecx
'0070886a    89bd34feffff            mov dword ptr [ebp+fffffe34], edi
'00708870    898d2cfeffff            mov dword ptr [ebp+fffffe2c], ecx
'00708876    89bd44feffff            mov dword ptr [ebp+fffffe44], edi
'0070887c    898d3cfeffff            mov dword ptr [ebp+fffffe3c], ecx
'00708882    89bd54feffff            mov dword ptr [ebp+fffffe54], edi
'00708888    898d4cfeffff            mov dword ptr [ebp+fffffe4c], ecx
'0070888e    89bd64feffff            mov dword ptr [ebp+fffffe64], edi
'00708894    898d5cfeffff            mov dword ptr [ebp+fffffe5c], ecx
'0070889a    89bd74feffff            mov dword ptr [ebp+fffffe74], edi
'007088a0    898d6cfeffff            mov dword ptr [ebp+fffffe6c], ecx
'007088a6    89bd84feffff            mov dword ptr [ebp+fffffe84], edi
'007088ac    898d7cfeffff            mov dword ptr [ebp+fffffe7c], ecx
'007088b2    8bbd74ffffff            mov edi, dword ptr [ebp+ffffff74]
'007088b8    899d74ffffff            mov dword ptr [ebp+ffffff74], ebx
'007088be    89bd64ffffff            mov dword ptr [ebp+ffffff64], edi
'007088c4    c7855cffffff09000000    mov dword ptr [ebp+ffffff5c], 00000009
'007088ce    8b7dc4                  mov edi, dword ptr [ebp-3c]
'007088d1    8b3f                    mov edi, dword ptr [edi]
'007088d3    83ec10                  sub esp, 10
'007088d6    8bdc                    mov ebx, esp
'007088d8    890b                    mov dword ptr [ebx], ecx
'007088da    8b8dd0fdffff            mov ecx, dword ptr [ebp+fffffdd0]
'007088e0    894b04                  mov dword ptr [ebx+04], ecx
'007088e3    894308                  mov dword ptr [ebx+08], eax
'007088e6    8b85d8fdffff            mov eax, dword ptr [ebp+fffffdd8]
'007088ec    89430c                  mov dword ptr [ebx+0c], eax
'007088ef    83ec10                  sub esp, 10
'007088f2    8bcc                    mov ecx, esp
'007088f4    8931                    mov dword ptr [ecx], esi
'007088f6    8b85e0fdffff            mov eax, dword ptr [ebp+fffffde0]
'007088fc    894104                  mov dword ptr [ecx+04], eax
'007088ff    895108                  mov dword ptr [ecx+08], edx
'00708902    8b95e8fdffff            mov edx, dword ptr [ebp+fffffde8]
'00708908    89510c                  mov dword ptr [ecx+0c], edx
'0070890b    83ec10                  sub esp, 10
'0070890e    8bc4                    mov eax, esp
'00708910    8bce                    mov ecx, esi
'00708912    8908                    mov dword ptr [eax], ecx
'00708914    8b95f0fdffff            mov edx, dword ptr [ebp+fffffdf0]
'0070891a    895004                  mov dword ptr [eax+04], edx
'0070891d    8b8df4fdffff            mov ecx, dword ptr [ebp+fffffdf4]
'00708923    894808                  mov dword ptr [eax+08], ecx
'00708926    8b95f8fdffff            mov edx, dword ptr [ebp+fffffdf8]
'0070892c    89500c                  mov dword ptr [eax+0c], edx
'0070892f    83ec10                  sub esp, 10
'00708932    8bc4                    mov eax, esp
'00708934    8bce                    mov ecx, esi
'00708936    8908                    mov dword ptr [eax], ecx
'00708938    8b9500feffff            mov edx, dword ptr [ebp+fffffe00]
'0070893e    895004                  mov dword ptr [eax+04], edx
'00708941    8b8d04feffff            mov ecx, dword ptr [ebp+fffffe04]
'00708947    894808                  mov dword ptr [eax+08], ecx
'0070894a    8b9508feffff            mov edx, dword ptr [ebp+fffffe08]
'00708950    89500c                  mov dword ptr [eax+0c], edx
'00708953    83ec10                  sub esp, 10
'00708956    8bc4                    mov eax, esp
'00708958    8bce                    mov ecx, esi
'0070895a    8908                    mov dword ptr [eax], ecx
'0070895c    8b9510feffff            mov edx, dword ptr [ebp+fffffe10]
'00708962    895004                  mov dword ptr [eax+04], edx
'00708965    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'0070896b    894808                  mov dword ptr [eax+08], ecx
'0070896e    8b9518feffff            mov edx, dword ptr [ebp+fffffe18]
'00708974    89500c                  mov dword ptr [eax+0c], edx
'00708977    83ec10                  sub esp, 10
'0070897a    8bc4                    mov eax, esp
'0070897c    8bce                    mov ecx, esi
'0070897e    8908                    mov dword ptr [eax], ecx
'00708980    8b9520feffff            mov edx, dword ptr [ebp+fffffe20]
'00708986    895004                  mov dword ptr [eax+04], edx
'00708989    8b8d24feffff            mov ecx, dword ptr [ebp+fffffe24]
'0070898f    894808                  mov dword ptr [eax+08], ecx
'00708992    8b9528feffff            mov edx, dword ptr [ebp+fffffe28]
'00708998    89500c                  mov dword ptr [eax+0c], edx
'0070899b    83ec10                  sub esp, 10
'0070899e    8bc4                    mov eax, esp
'007089a0    8bce                    mov ecx, esi
'007089a2    8908                    mov dword ptr [eax], ecx
'007089a4    8b9530feffff            mov edx, dword ptr [ebp+fffffe30]
'007089aa    895004                  mov dword ptr [eax+04], edx
'007089ad    8b8d34feffff            mov ecx, dword ptr [ebp+fffffe34]
'007089b3    894808                  mov dword ptr [eax+08], ecx
'007089b6    8b9538feffff            mov edx, dword ptr [ebp+fffffe38]
'007089bc    89500c                  mov dword ptr [eax+0c], edx
'007089bf    83ec10                  sub esp, 10
'007089c2    8bc4                    mov eax, esp
'007089c4    8bce                    mov ecx, esi
'007089c6    8908                    mov dword ptr [eax], ecx
'007089c8    8b9540feffff            mov edx, dword ptr [ebp+fffffe40]
'007089ce    895004                  mov dword ptr [eax+04], edx
'007089d1    8b8d44feffff            mov ecx, dword ptr [ebp+fffffe44]
'007089d7    894808                  mov dword ptr [eax+08], ecx
'007089da    8b9548feffff            mov edx, dword ptr [ebp+fffffe48]
'007089e0    89500c                  mov dword ptr [eax+0c], edx
'007089e3    83ec10                  sub esp, 10
'007089e6    8bc4                    mov eax, esp
'007089e8    8bce                    mov ecx, esi
'007089ea    8908                    mov dword ptr [eax], ecx
'007089ec    8b9550feffff            mov edx, dword ptr [ebp+fffffe50]
'007089f2    895004                  mov dword ptr [eax+04], edx
'007089f5    8b8d54feffff            mov ecx, dword ptr [ebp+fffffe54]
'007089fb    894808                  mov dword ptr [eax+08], ecx
'007089fe    8b9558feffff            mov edx, dword ptr [ebp+fffffe58]
'00708a04    89500c                  mov dword ptr [eax+0c], edx
'00708a07    83ec10                  sub esp, 10
'00708a0a    8bc4                    mov eax, esp
'00708a0c    8bce                    mov ecx, esi
'00708a0e    8908                    mov dword ptr [eax], ecx
'00708a10    8b9560feffff            mov edx, dword ptr [ebp+fffffe60]
'00708a16    895004                  mov dword ptr [eax+04], edx
'00708a19    8b8d64feffff            mov ecx, dword ptr [ebp+fffffe64]
'00708a1f    894808                  mov dword ptr [eax+08], ecx
'00708a22    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'00708a28    89500c                  mov dword ptr [eax+0c], edx
'00708a2b    83ec10                  sub esp, 10
'00708a2e    8bc4                    mov eax, esp
'00708a30    8bce                    mov ecx, esi
'00708a32    8908                    mov dword ptr [eax], ecx
'00708a34    8b9570feffff            mov edx, dword ptr [ebp+fffffe70]
'00708a3a    895004                  mov dword ptr [eax+04], edx
'00708a3d    8b8d74feffff            mov ecx, dword ptr [ebp+fffffe74]
'00708a43    894808                  mov dword ptr [eax+08], ecx
'00708a46    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'00708a4c    89500c                  mov dword ptr [eax+0c], edx
'00708a4f    83ec10                  sub esp, 10
'00708a52    8bc4                    mov eax, esp
'00708a54    8bce                    mov ecx, esi
'00708a56    8908                    mov dword ptr [eax], ecx
'00708a58    8b9580feffff            mov edx, dword ptr [ebp+fffffe80]
'00708a5e    895004                  mov dword ptr [eax+04], edx
'00708a61    8b8d84feffff            mov ecx, dword ptr [ebp+fffffe84]
'00708a67    894808                  mov dword ptr [eax+08], ecx
'00708a6a    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'00708a70    89500c                  mov dword ptr [eax+0c], edx
'00708a73    83ec10                  sub esp, 10
'00708a76    8bc4                    mov eax, esp
'00708a78    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'00708a7e    8908                    mov dword ptr [eax], ecx
'00708a80    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'00708a86    895004                  mov dword ptr [eax+04], edx
'00708a89    8b8d64ffffff            mov ecx, dword ptr [ebp+ffffff64]
'00708a8f    894808                  mov dword ptr [eax+08], ecx
'00708a92    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'00708a98    89500c                  mov dword ptr [eax+0c], edx
'00708a9b    68209e4300              push 00439e20
'00708aa0    8b45c4                  mov eax, dword ptr [ebp-3c]
'00708aa3    50                      push eax
'00708aa4    ff97f4000000            call dword ptr [edi+000000f4]
'00708aaa    dbe2                    fnclex
'00708aac    85c0                    test ax, ax
'00708aae    7d15                    jge 708ac5
'00708ab0    68f4000000              push 000000f4
'00708ab5    6830314300              push 00433130
'00708aba    8b4dc4                  mov ecx, dword ptr [ebp-3c]
'00708abd    51                      push ecx
'00708abe    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00708abf    ff1580104000            call dword ptr [00401080]
'00708ac5    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeObj"
'00708acb    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_87)
'00708ad1    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]

' *** Reference to "__vbaFreeVar"
'00708ad7    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_88)
'00708add    8b45a4                  mov eax, dword ptr [ebp-5c]
'00708ae0    8b10                    mov edx, dword ptr [eax]
'00708ae2    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'00708ae8    51                      push ecx
'00708ae9    50                      push eax
'00708aea    ff92b4000000            call dword ptr [edx+000000b4]
'00708af0    dbe2                    fnclex
'00708af2    85c0                    test ax, ax
'00708af4    7d15                    jge 708b0b
'00708af6    68b4000000              push 000000b4
'00708afb    6830314300              push 00433130
'00708b00    8b55a4                  mov edx, dword ptr [ebp-5c]
'00708b03    52                      push edx
'00708b04    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00708b05    ff1580104000            call dword ptr [00401080]
'00708b0b    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'00708b11    8bf0                    mov esi, eax
'00708b13    b938ab4300              mov ecx, 0043ab38
'00708b18    898d94feffff            mov dword ptr [ebp+fffffe94], ecx
'00708b1e    ba08000000              mov edx, 00000008
'00708b23    89958cfeffff            mov dword ptr [ebp+fffffe8c], edx
'00708b29    8b38                    mov edi, dword ptr [eax]
'00708b2b    8d9d74ffffff            lea ebx, dword ptr [ebp+ffffff74]
'00708b31    53                      push ebx
'00708b32    83ec10                  sub esp, 10
'00708b35    8bdc                    mov ebx, esp
'00708b37    8913                    mov dword ptr [ebx], edx
'00708b39    8b9590feffff            mov edx, dword ptr [ebp+fffffe90]
'00708b3f    895304                  mov dword ptr [ebx+04], edx
'00708b42    894b08                  mov dword ptr [ebx+08], ecx
'00708b45    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'00708b4b    894b0c                  mov dword ptr [ebx+0c], ecx
'00708b4e    50                      push eax
'00708b4f    ff5730                  call dword ptr [edi+30]
'00708b52    dbe2                    fnclex
'00708b54    85c0                    test ax, ax
'00708b56    7d0f                    jge 708b67
'00708b58    6a30                    push 30
'00708b5a    68d8304300              push 004330d8
'00708b5f    56                      push esi
'00708b60    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00708b61    ff1580104000            call dword ptr [00401080]
'00708b67    b804000280              mov eax, 80020004
'00708b6c    b90a000000              mov ecx, 0000000a
'00708b71    8bd0                    mov edx, eax
'00708b73    8995e4fdffff            mov dword ptr [ebp+fffffde4], edx
'00708b79    8bf1                    mov esi, ecx
'00708b7b    89b5dcfdffff            mov dword ptr [ebp+fffffddc], esi
'00708b81    8985f4fdffff            mov dword ptr [ebp+fffffdf4], eax
'00708b87    898decfdffff            mov dword ptr [ebp+fffffdec], ecx
'00708b8d    898504feffff            mov dword ptr [ebp+fffffe04], eax
'00708b93    898dfcfdffff            mov dword ptr [ebp+fffffdfc], ecx
'00708b99    898514feffff            mov dword ptr [ebp+fffffe14], eax
'00708b9f    898d0cfeffff            mov dword ptr [ebp+fffffe0c], ecx
'00708ba5    898524feffff            mov dword ptr [ebp+fffffe24], eax
'00708bab    898d1cfeffff            mov dword ptr [ebp+fffffe1c], ecx
'00708bb1    898534feffff            mov dword ptr [ebp+fffffe34], eax
'00708bb7    898d2cfeffff            mov dword ptr [ebp+fffffe2c], ecx
'00708bbd    898544feffff            mov dword ptr [ebp+fffffe44], eax
'00708bc3    898d3cfeffff            mov dword ptr [ebp+fffffe3c], ecx
'00708bc9    898554feffff            mov dword ptr [ebp+fffffe54], eax
'00708bcf    898d4cfeffff            mov dword ptr [ebp+fffffe4c], ecx
'00708bd5    898564feffff            mov dword ptr [ebp+fffffe64], eax
'00708bdb    898d5cfeffff            mov dword ptr [ebp+fffffe5c], ecx
'00708be1    898574feffff            mov dword ptr [ebp+fffffe74], eax
'00708be7    898d6cfeffff            mov dword ptr [ebp+fffffe6c], ecx
'00708bed    898584feffff            mov dword ptr [ebp+fffffe84], eax
'00708bf3    898d7cfeffff            mov dword ptr [ebp+fffffe7c], ecx
'00708bf9    8bbd74ffffff            mov edi, dword ptr [ebp+ffffff74]
'00708bff    c78574ffffff00000000    mov dword ptr [ebp+ffffff74], 00000000
'00708c09    89bd64ffffff            mov dword ptr [ebp+ffffff64], edi
'00708c0f    c7855cffffff09000000    mov dword ptr [ebp+ffffff5c], 00000009
'00708c19    8b7db8                  mov edi, dword ptr [ebp-48]
'00708c1c    8b3f                    mov edi, dword ptr [edi]
'00708c1e    83ec10                  sub esp, 10
'00708c21    8bdc                    mov ebx, esp
'00708c23    890b                    mov dword ptr [ebx], ecx
'00708c25    8b8dd0fdffff            mov ecx, dword ptr [ebp+fffffdd0]
'00708c2b    894b04                  mov dword ptr [ebx+04], ecx
'00708c2e    894308                  mov dword ptr [ebx+08], eax
'00708c31    8b85d8fdffff            mov eax, dword ptr [ebp+fffffdd8]
'00708c37    89430c                  mov dword ptr [ebx+0c], eax
'00708c3a    83ec10                  sub esp, 10
'00708c3d    8bcc                    mov ecx, esp
'00708c3f    8931                    mov dword ptr [ecx], esi
'00708c41    8b85e0fdffff            mov eax, dword ptr [ebp+fffffde0]
'00708c47    894104                  mov dword ptr [ecx+04], eax
'00708c4a    895108                  mov dword ptr [ecx+08], edx
'00708c4d    8b95e8fdffff            mov edx, dword ptr [ebp+fffffde8]
'00708c53    89510c                  mov dword ptr [ecx+0c], edx
'00708c56    83ec10                  sub esp, 10
'00708c59    8bc4                    mov eax, esp
'00708c5b    8bce                    mov ecx, esi
'00708c5d    8908                    mov dword ptr [eax], ecx
'00708c5f    8b95f0fdffff            mov edx, dword ptr [ebp+fffffdf0]
'00708c65    895004                  mov dword ptr [eax+04], edx
'00708c68    8b8df4fdffff            mov ecx, dword ptr [ebp+fffffdf4]
'00708c6e    894808                  mov dword ptr [eax+08], ecx
'00708c71    8b95f8fdffff            mov edx, dword ptr [ebp+fffffdf8]
'00708c77    89500c                  mov dword ptr [eax+0c], edx
'00708c7a    83ec10                  sub esp, 10
'00708c7d    8bc4                    mov eax, esp
'00708c7f    8bce                    mov ecx, esi
'00708c81    8908                    mov dword ptr [eax], ecx
'00708c83    8b9500feffff            mov edx, dword ptr [ebp+fffffe00]
'00708c89    895004                  mov dword ptr [eax+04], edx
'00708c8c    8b8d04feffff            mov ecx, dword ptr [ebp+fffffe04]
'00708c92    894808                  mov dword ptr [eax+08], ecx
'00708c95    8b9508feffff            mov edx, dword ptr [ebp+fffffe08]
'00708c9b    89500c                  mov dword ptr [eax+0c], edx
'00708c9e    83ec10                  sub esp, 10
'00708ca1    8bc4                    mov eax, esp
'00708ca3    8bce                    mov ecx, esi
'00708ca5    8908                    mov dword ptr [eax], ecx
'00708ca7    8b9510feffff            mov edx, dword ptr [ebp+fffffe10]
'00708cad    895004                  mov dword ptr [eax+04], edx
'00708cb0    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'00708cb6    894808                  mov dword ptr [eax+08], ecx
'00708cb9    8b9518feffff            mov edx, dword ptr [ebp+fffffe18]
'00708cbf    89500c                  mov dword ptr [eax+0c], edx
'00708cc2    83ec10                  sub esp, 10
'00708cc5    8bc4                    mov eax, esp
'00708cc7    8bce                    mov ecx, esi
'00708cc9    8908                    mov dword ptr [eax], ecx
'00708ccb    8b9520feffff            mov edx, dword ptr [ebp+fffffe20]
'00708cd1    895004                  mov dword ptr [eax+04], edx
'00708cd4    8b8d24feffff            mov ecx, dword ptr [ebp+fffffe24]
'00708cda    894808                  mov dword ptr [eax+08], ecx
'00708cdd    8b9528feffff            mov edx, dword ptr [ebp+fffffe28]
'00708ce3    89500c                  mov dword ptr [eax+0c], edx
'00708ce6    83ec10                  sub esp, 10
'00708ce9    8bc4                    mov eax, esp
'00708ceb    8bce                    mov ecx, esi
'00708ced    8908                    mov dword ptr [eax], ecx
'00708cef    8b9530feffff            mov edx, dword ptr [ebp+fffffe30]
'00708cf5    895004                  mov dword ptr [eax+04], edx
'00708cf8    8b8d34feffff            mov ecx, dword ptr [ebp+fffffe34]
'00708cfe    894808                  mov dword ptr [eax+08], ecx
'00708d01    8b9538feffff            mov edx, dword ptr [ebp+fffffe38]
'00708d07    89500c                  mov dword ptr [eax+0c], edx
'00708d0a    83ec10                  sub esp, 10
'00708d0d    8bc4                    mov eax, esp
'00708d0f    8bce                    mov ecx, esi
'00708d11    8908                    mov dword ptr [eax], ecx
'00708d13    8b9540feffff            mov edx, dword ptr [ebp+fffffe40]
'00708d19    895004                  mov dword ptr [eax+04], edx
'00708d1c    8b8d44feffff            mov ecx, dword ptr [ebp+fffffe44]
'00708d22    894808                  mov dword ptr [eax+08], ecx
'00708d25    8b9548feffff            mov edx, dword ptr [ebp+fffffe48]
'00708d2b    89500c                  mov dword ptr [eax+0c], edx
'00708d2e    83ec10                  sub esp, 10
'00708d31    8bc4                    mov eax, esp
'00708d33    8bce                    mov ecx, esi
'00708d35    8908                    mov dword ptr [eax], ecx
'00708d37    8b9550feffff            mov edx, dword ptr [ebp+fffffe50]
'00708d3d    895004                  mov dword ptr [eax+04], edx
'00708d40    8b8d54feffff            mov ecx, dword ptr [ebp+fffffe54]
'00708d46    894808                  mov dword ptr [eax+08], ecx
'00708d49    8b9558feffff            mov edx, dword ptr [ebp+fffffe58]
'00708d4f    89500c                  mov dword ptr [eax+0c], edx
'00708d52    83ec10                  sub esp, 10
'00708d55    8bc4                    mov eax, esp
'00708d57    8bce                    mov ecx, esi
'00708d59    8908                    mov dword ptr [eax], ecx
'00708d5b    8b9560feffff            mov edx, dword ptr [ebp+fffffe60]
'00708d61    895004                  mov dword ptr [eax+04], edx
'00708d64    8b8d64feffff            mov ecx, dword ptr [ebp+fffffe64]
'00708d6a    894808                  mov dword ptr [eax+08], ecx
'00708d6d    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'00708d73    89500c                  mov dword ptr [eax+0c], edx
'00708d76    83ec10                  sub esp, 10
'00708d79    8bc4                    mov eax, esp
'00708d7b    8bce                    mov ecx, esi
'00708d7d    8908                    mov dword ptr [eax], ecx
'00708d7f    8b9570feffff            mov edx, dword ptr [ebp+fffffe70]
'00708d85    895004                  mov dword ptr [eax+04], edx
'00708d88    8b8d74feffff            mov ecx, dword ptr [ebp+fffffe74]
'00708d8e    894808                  mov dword ptr [eax+08], ecx
'00708d91    8b9578feffff            mov edx, dword ptr [ebp+fffffe78]
'00708d97    89500c                  mov dword ptr [eax+0c], edx
'00708d9a    83ec10                  sub esp, 10
'00708d9d    8bc4                    mov eax, esp
'00708d9f    8bce                    mov ecx, esi
'00708da1    8908                    mov dword ptr [eax], ecx
'00708da3    8b9580feffff            mov edx, dword ptr [ebp+fffffe80]
'00708da9    895004                  mov dword ptr [eax+04], edx
'00708dac    8b8d84feffff            mov ecx, dword ptr [ebp+fffffe84]
'00708db2    894808                  mov dword ptr [eax+08], ecx
'00708db5    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'00708dbb    89500c                  mov dword ptr [eax+0c], edx
'00708dbe    83ec10                  sub esp, 10
'00708dc1    8bc4                    mov eax, esp
'00708dc3    8b8d5cffffff            mov ecx, dword ptr [ebp+ffffff5c]
'00708dc9    8908                    mov dword ptr [eax], ecx
'00708dcb    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'00708dd1    895004                  mov dword ptr [eax+04], edx
'00708dd4    8b8d64ffffff            mov ecx, dword ptr [ebp+ffffff64]
'00708dda    894808                  mov dword ptr [eax+08], ecx
'00708ddd    8b9568ffffff            mov edx, dword ptr [ebp+ffffff68]
'00708de3    89500c                  mov dword ptr [eax+0c], edx
'00708de6    68209e4300              push 00439e20
'00708deb    8b45b8                  mov eax, dword ptr [ebp-48]
'00708dee    50                      push eax
'00708def    ff97f4000000            call dword ptr [edi+000000f4]
'00708df5    dbe2                    fnclex
'00708df7    85c0                    test ax, ax
'00708df9    7d15                    jge 708e10
'00708dfb    68f4000000              push 000000f4
'00708e00    6830314300              push 00433130
'00708e05    8b4db8                  mov ecx, dword ptr [ebp-48]
'00708e08    51                      push ecx
'00708e09    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00708e0a    ff1580104000            call dword ptr [00401080]
'00708e10    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeObj"
'00708e16    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_87)
'00708e1c    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]

' *** Reference to "__vbaFreeVar"
'00708e22    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_88)
'00708e28    8b45a4                  mov eax, dword ptr [ebp-5c]
'00708e2b    8b10                    mov edx, dword ptr [eax]
'00708e2d    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'00708e33    51                      push ecx
'00708e34    50                      push eax
'00708e35    ff92b4000000            call dword ptr [edx+000000b4]
'00708e3b    dbe2                    fnclex
'00708e3d    85c0                    test ax, ax
'00708e3f    7d15                    jge 708e56
'00708e41    68b4000000              push 000000b4
'00708e46    6830314300              push 00433130
'00708e4b    8b55a4                  mov edx, dword ptr [ebp-5c]
'00708e4e    52                      push edx
'00708e4f    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00708e50    ff1580104000            call dword ptr [00401080]
'00708e56    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'00708e5c    8bf0                    mov esi, eax
'00708e5e    b964b14300              mov ecx, 0043b164
'00708e63    898d94feffff            mov dword ptr [ebp+fffffe94], ecx
'00708e69    ba08000000              mov edx, 00000008
'00708e6e    89958cfeffff            mov dword ptr [ebp+fffffe8c], edx
'00708e74    8b38                    mov edi, dword ptr [eax]
'00708e76    8d9d74ffffff            lea ebx, dword ptr [ebp+ffffff74]
'00708e7c    53                      push ebx
'00708e7d    83ec10                  sub esp, 10
'00708e80    8bdc                    mov ebx, esp
'00708e82    8913                    mov dword ptr [ebx], edx
'00708e84    8b9590feffff            mov edx, dword ptr [ebp+fffffe90]
'00708e8a    895304                  mov dword ptr [ebx+04], edx
'00708e8d    894b08                  mov dword ptr [ebx+08], ecx
'00708e90    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'00708e96    894b0c                  mov dword ptr [ebx+0c], ecx
'00708e99    50                      push eax
'00708e9a    ff5730                  call dword ptr [edi+30]
'00708e9d    dbe2                    fnclex
'00708e9f    85c0                    test ax, ax
'00708ea1    7d0f                    jge 708eb2
'00708ea3    6a30                    push 30
'00708ea5    68d8304300              push 004330d8
'00708eaa    56                      push esi
'00708eab    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00708eac    ff1580104000            call dword ptr [00401080]
'00708eb2    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'00708eb8    8bf0                    mov esi, eax
'00708eba    8b10                    mov edx, dword ptr [eax]
'00708ebc    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]
'00708ec2    51                      push ecx
'00708ec3    50                      push eax
'00708ec4    ff5244                  call dword ptr [edx+44]
'00708ec7    dbe2                    fnclex
'00708ec9    85c0                    test ax, ax
'00708ecb    7d0f                    jge 708edc
'00708ecd    6a44                    push 44
'00708ecf    6880304300              push 00433080
'00708ed4    56                      push esi
'00708ed5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'00708ed6    ff1580104000            call dword ptr [00401080]
'00708edc    8d955cffffff            lea edx, dword ptr [ebp+ffffff5c]
'00708ee2    52                      push edx

' *** Reference to "__vbaStrVarMove"
'00708ee3    ff153c104000            call dword ptr [0040103c]
'00708ee9    8bd0                    mov edx, eax
'00708eeb    8d4d94                  lea ecx, dword ptr [ebp-6c]

' *** Reference to "__vbaStrMove"
'00708eee    8b35d0124000            mov esi, dword ptr [004012d0]
'00708ef4    ffd6                    call esi
'00708ef6    8d4594                  lea eax, dword ptr [ebp-6c]
'00708ef9    50                      push eax
'00708efa    e8f1acdeff              call 4f3bf0
Call sub_4F3BF0()
'00708eff    8bd0                    mov edx, eax
'00708f01    8d4d84                  lea ecx, dword ptr [ebp-7c]
'00708f04    ffd6                    call esi
'00708f06    b804000280              mov eax, 80020004
'00708f0b    898564feffff            mov dword ptr [ebp+fffffe64], eax
'00708f11    b90a000000              mov ecx, 0000000a
'00708f16    898d5cfeffff            mov dword ptr [ebp+fffffe5c], ecx
'00708f1c    8bf0                    mov esi, eax
'00708f1e    89b574feffff            mov dword ptr [ebp+fffffe74], esi
'00708f24    8bf9                    mov edi, ecx
'00708f26    89bd6cfeffff            mov dword ptr [ebp+fffffe6c], edi
'00708f2c    c78584feffff02000000    mov dword ptr [ebp+fffffe84], 00000002
'00708f36    c7857cfeffff03000000    mov dword ptr [ebp+fffffe7c], 00000003
'00708f40    8b5584                  mov edx, dword ptr [ebp-7c]
'00708f43    899554fdffff            mov dword ptr [ebp+fffffd54], edx
'00708f49    c7458400000000          mov dword ptr [ebp-7c], 00000000
'00708f50    8b1548b07200            mov edx, dword ptr [0072b048]
'00708f56    8b1a                    mov ebx, dword ptr [edx]
'00708f58    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'00708f5e    52                      push edx
'00708f5f    83ec10                  sub esp, 10
'00708f62    8bd4                    mov edx, esp
'00708f64    890a                    mov dword ptr [edx], ecx
'00708f66    8b8d60feffff            mov ecx, dword ptr [ebp+fffffe60]
'00708f6c    894a04                  mov dword ptr [edx+04], ecx
'00708f6f    894208                  mov dword ptr [edx+08], eax
'00708f72    8b8568feffff            mov eax, dword ptr [ebp+fffffe68]
'00708f78    89420c                  mov dword ptr [edx+0c], eax
'00708f7b    83ec10                  sub esp, 10
'00708f7e    8bcc                    mov ecx, esp
'00708f80    8939                    mov dword ptr [ecx], edi
'00708f82    8b9570feffff            mov edx, dword ptr [ebp+fffffe70]
'00708f88    895104                  mov dword ptr [ecx+04], edx
'00708f8b    897108                  mov dword ptr [ecx+08], esi
'00708f8e    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'00708f94    89410c                  mov dword ptr [ecx+0c], eax
'00708f97    83ec10                  sub esp, 10
'00708f9a    8bcc                    mov ecx, esp
'00708f9c    8b957cfeffff            mov edx, dword ptr [ebp+fffffe7c]
'00708fa2    8911                    mov dword ptr [ecx], edx
'00708fa4    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'00708faa    894104                  mov dword ptr [ecx+04], eax
'00708fad    8b9584feffff            mov edx, dword ptr [ebp+fffffe84]
'00708fb3    895108                  mov dword ptr [ecx+08], edx
'00708fb6    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'00708fbc    89410c                  mov dword ptr [ecx+0c], eax
'00708fbf    6878624500              push 00456278
'00708fc4    8b9554fdffff            mov edx, dword ptr [ebp+fffffd54]
'00708fca    8d4d90                  lea ecx, dword ptr [ebp-70]

' *** Reference to "__vbaStrMove"
'00708fcd    8b3dd0124000            mov edi, dword ptr [004012d0]
'00708fd3    ffd7                    call edi
'00708fd5    50                      push eax

' *** Reference to "__vbaStrCat"
'00708fd6    8b3570104000            mov esi, dword ptr [00401070]
'00708fdc    ffd6                    call esi
var_121 = ("select dons from PersonnageDons where nom='") & (-4512)
'00708fde    8bd0                    mov edx, eax
'00708fe0    8d4d8c                  lea ecx, dword ptr [ebp-74]
'00708fe3    ffd7                    call edi
'00708fe5    50                      push eax
'00708fe6    6854a44300              push 0043a454
'00708feb    ffd6                    call esi
var_13 = (var_121) & ("'")
'00708fed    8bd0                    mov edx, eax
'00708fef    8d4d88                  lea ecx, dword ptr [ebp-78]
'00708ff2    ffd7                    call edi
'00708ff4    50                      push eax
'00708ff5    8b0d48b07200            mov ecx, dword ptr [0072b048]
'00708ffb    51                      push ecx
'00708ffc    ff93bc000000            call dword ptr [ebx+000000bc]
'00709002    dbe2                    fnclex
'00709004    85c0                    test ax, ax
'00709006    0f8db8030000            jge 7093c4
'0070900c    68bc000000              push 000000bc
'00709011    68301f4300              push 00431f30
'00709016    8b1548b07200            mov edx, dword ptr [0072b048]
'0070901c    52                      push edx
'0070901d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070901e    8b1d80104000            mov ebx, dword ptr [00401080]
'00709024    ffd3                    call ebx
'00709026    e99f030000              jmp 7093ca

' *** Reference to "rtcErrObj"
'0070902b    8b1d6c124000            mov ebx, dword ptr [0040126c]
'00709031    ffd3                    call ebx
'00709033    50                      push eax
'00709034    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0070903a    51                      push ecx

' *** Reference to "__vbaObjSet"
'0070903b    8b3db4104000            mov edi, dword ptr [004010b4]
'00709041    ffd7                    call edi
Set var_87 = Err
'00709043    8bf0                    mov esi, eax
'00709045    8b16                    mov edx, dword ptr [esi]
'00709047    8d4594                  lea eax, dword ptr [ebp-6c]
'0070904a    50                      push eax
'0070904b    56                      push esi
'0070904c    ff522c                  call dword ptr [edx+2c]
var_121 = var_87.Description
'0070904f    dbe2                    fnclex
'00709051    85c0                    test ax, ax
'00709053    7d0f                    jge 709064
'00709055    6a2c                    push 2c
'00709057    685c084300              push 0043085c
'0070905c    56                      push esi
'0070905d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070905e    ff1580104000            call dword ptr [00401080]
'00709064    ffd3                    call ebx
'00709066    50                      push eax
'00709067    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'0070906d    51                      push ecx
'0070906e    ffd7                    call edi
Set var_91 = Err
'00709070    8bf0                    mov esi, eax
'00709072    8b16                    mov edx, dword ptr [esi]
'00709074    8d8594fdffff            lea eax, dword ptr [ebp+fffffd94]
'0070907a    50                      push eax
'0070907b    56                      push esi
'0070907c    ff521c                  call dword ptr [edx+1c]
var_500 = var_91.Number
'0070907f    dbe2                    fnclex
'00709081    85c0                    test ax, ax
'00709083    7d0f                    jge 709094
'00709085    6a1c                    push 1c
'00709087    685c084300              push 0043085c
'0070908c    56                      push esi
'0070908d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070908e    ff1580104000            call dword ptr [00401080]
'00709094    b904000280              mov ecx, 80020004
'00709099    898d44ffffff            mov dword ptr [ebp+ffffff44], ecx
'0070909f    b80a000000              mov eax, 0000000a
'007090a4    89853cffffff            mov dword ptr [ebp+ffffff3c], eax
'007090aa    898d54ffffff            mov dword ptr [ebp+ffffff54], ecx
'007090b0    89854cffffff            mov dword ptr [ebp+ffffff4c], eax
'007090b6    c78594feffff24b07200    mov dword ptr [ebp+fffffe94], 0072b024
'007090c0    c7858cfeffff08400000    mov dword ptr [ebp+fffffe8c], 00004008
'007090ca    6814084300              push 00430814
'007090cf    8b4d94                  mov ecx, dword ptr [ebp-6c]
'007090d2    51                      push ecx

' *** Reference to "__vbaStrCat"
'007090d3    8b3570104000            mov esi, dword ptr [00401070]
'007090d9    ffd6                    call esi
var_11 = ("L'erreur suivante s'est produite : ") & (var_121)
'007090db    8bd0                    mov edx, eax
'007090dd    8d4d90                  lea ecx, dword ptr [ebp-70]

' *** Reference to "__vbaStrMove"
'007090e0    8b3dd0124000            mov edi, dword ptr [004012d0]
'007090e6    ffd7                    call edi
'007090e8    50                      push eax
'007090e9    6870084300              push 00430870
'007090ee    ffd6                    call esi
var_128 = (var_11) & (vbCrLf)
'007090f0    8bd0                    mov edx, eax
'007090f2    8d4d8c                  lea ecx, dword ptr [ebp-74]
'007090f5    ffd7                    call edi
'007090f7    50                      push eax
'007090f8    6870084300              push 00430870
'007090fd    ffd6                    call esi
var_17 = (var_128) & (vbCrLf)
'007090ff    8bd0                    mov edx, eax
'00709101    8d4d88                  lea ecx, dword ptr [ebp-78]
'00709104    ffd7                    call edi
'00709106    50                      push eax
'00709107    6880084300              push 00430880
'0070910c    ffd6                    call esi
var_129 = (var_17) & (" numéro d'erreur (")
'0070910e    8bd0                    mov edx, eax
'00709110    8d4d84                  lea ecx, dword ptr [ebp-7c]
'00709113    ffd7                    call edi
'00709115    50                      push eax
'00709116    8b9594fdffff            mov edx, dword ptr [ebp+fffffd94]
'0070911c    52                      push edx

' *** Reference to "__vbaStrI4"
'0070911d    ff1520104000            call dword ptr [00401020]
'00709123    8bd0                    mov edx, eax
'00709125    8d4d80                  lea ecx, dword ptr [ebp-80]
'00709128    ffd7                    call edi
'0070912a    50                      push eax
'0070912b    ffd6                    call esi
var_285 = (var_129) & (CStr(var_500))
'0070912d    8bd0                    mov edx, eax
'0070912f    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'00709135    ffd7                    call edi
'00709137    50                      push eax
'00709138    68ac084300              push 004308ac
'0070913d    ffd6                    call esi
var_54 = (var_285) & (" )")
'0070913f    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'00709145    bf08000000              mov edi, 00000008
'0070914a    89bd5cffffff            mov dword ptr [ebp+ffffff5c], edi
'00709150    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'00709156    50                      push eax
'00709157    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'0070915d    51                      push ecx
'0070915e    8d958cfeffff            lea edx, dword ptr [ebp+fffffe8c]
'00709164    52                      push edx
'00709165    6a10                    push 10
'00709167    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070916d    50                      push eax

' *** Reference to "rtcMsgBox"
'0070916e    ff15b8104000            call dword ptr [004010b8]
var_139 = MsgBox(var_54, 16, vbNullString)
'00709174    8d8d7cffffff            lea ecx, dword ptr [ebp+ffffff7c]
'0070917a    51                      push ecx
'0070917b    8d5580                  lea edx, dword ptr [ebp-80]
'0070917e    52                      push edx
'0070917f    8d4584                  lea eax, dword ptr [ebp-7c]
'00709182    50                      push eax
'00709183    8d4d88                  lea ecx, dword ptr [ebp-78]
'00709186    51                      push ecx
'00709187    8d558c                  lea edx, dword ptr [ebp-74]
'0070918a    52                      push edx
'0070918b    8d4590                  lea eax, dword ptr [ebp-70]
'0070918e    50                      push eax
'0070918f    8d4d94                  lea ecx, dword ptr [ebp-6c]
'00709192    51                      push ecx
'00709193    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'00709195    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4512, -4536, -4540, -4544, -4548, -4552, -4556)
'0070919b    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'007091a1    52                      push edx
'007091a2    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'007091a8    50                      push eax
'007091a9    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'007091ab    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_91)
'007091b1    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'007091b7    51                      push ecx
'007091b8    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'007091be    52                      push edx
'007091bf    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'007091c5    50                      push eax
'007091c6    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'007091c8    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_88, var_89, var_90)
'007091ce    83c43c                  add esp, 3c
'007091d1    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]
'007091d7    51                      push ecx

' *** Reference to "rtcGetPresentDate"
'007091d8    ff15f4124000            call dword ptr [004012f4]
var_54 = Now()
'007091de    c78594feffffb8084300    mov dword ptr [ebp+fffffe94], 004308b8
'007091e8    89bd8cfeffff            mov dword ptr [ebp+fffffe8c], edi
'007091ee    8d958cfeffff            lea edx, dword ptr [ebp+fffffe8c]
'007091f4    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]

' *** Reference to "__vbaVarDup"
'007091fa    ff1598124000            call dword ptr [00401298]
var_89 = ("dd/MM/yyyy hh:mm:ss")
'00709200    6a01                    push 01
'00709202    6a01                    push 01
'00709204    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'0070920a    52                      push edx
'0070920b    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'00709211    50                      push eax
'00709212    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'00709218    51                      push ecx

' *** Reference to "rtcVarFromFormatVar"
'00709219    ff1574104000            call dword ptr [00401074]
'0070921f    ffd3                    call ebx
'00709221    50                      push eax
'00709222    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'00709228    52                      push edx

' *** Reference to "__vbaObjSet"
'00709229    ff15b4104000            call dword ptr [004010b4]
Set var_87 = Err
'0070922f    8bf0                    mov esi, eax
'00709231    8b06                    mov eax, dword ptr [esi]
'00709233    8d8d94fdffff            lea ecx, dword ptr [ebp+fffffd94]
'00709239    51                      push ecx
'0070923a    56                      push esi
'0070923b    ff501c                  call dword ptr [eax+1c]
var_500 = var_87.Number
'0070923e    dbe2                    fnclex
'00709240    85c0                    test ax, ax
'00709242    7d0f                    jge 709253
'00709244    6a1c                    push 1c
'00709246    685c084300              push 0043085c
'0070924b    56                      push esi
'0070924c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070924d    ff1580104000            call dword ptr [00401080]
'00709253    ffd3                    call ebx
'00709255    50                      push eax
'00709256    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'0070925c    52                      push edx

' *** Reference to "__vbaObjSet"
'0070925d    ff15b4104000            call dword ptr [004010b4]
Set var_91 = Err
'00709263    8bf0                    mov esi, eax
'00709265    8b06                    mov eax, dword ptr [esi]
'00709267    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0070926a    51                      push ecx
'0070926b    56                      push esi
'0070926c    ff502c                  call dword ptr [eax+2c]
var_121 = var_91.Description
'0070926f    dbe2                    fnclex
'00709271    85c0                    test ax, ax
'00709273    7d0f                    jge 709284
'00709275    6a2c                    push 2c
'00709277    685c084300              push 0043085c
'0070927c    56                      push esi
'0070927d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070927e    ff1580104000            call dword ptr [00401080]
'00709284    c78584feffffe4084300    mov dword ptr [ebp+fffffe84], 004308e4
'0070928e    89bd7cfeffff            mov dword ptr [ebp+fffffe7c], edi
'00709294    8b9594fdffff            mov edx, dword ptr [ebp+fffffd94]
'0070929a    899574feffff            mov dword ptr [ebp+fffffe74], edx
'007092a0    c7856cfeffff03000000    mov dword ptr [ebp+fffffe6c], 00000003
'007092aa    c78564feffff08094300    mov dword ptr [ebp+fffffe64], 00430908
'007092b4    89bd5cfeffff            mov dword ptr [ebp+fffffe5c], edi
'007092ba    8b4594                  mov eax, dword ptr [ebp-6c]
'007092bd    c7459400000000          mov dword ptr [ebp-6c], 00000000
'007092c4    898504ffffff            mov dword ptr [ebp+ffffff04], eax
'007092ca    89bdfcfeffff            mov dword ptr [ebp+fffffefc], edi
'007092d0    c78554feffff50dd4400    mov dword ptr [ebp+fffffe54], 0044dd50
'007092da    89bd4cfeffff            mov dword ptr [ebp+fffffe4c], edi
'007092e0    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'007092e6    50                      push eax
'007092e7    8d8d7cfeffff            lea ecx, dword ptr [ebp+fffffe7c]
'007092ed    51                      push ecx
'007092ee    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'007092f4    52                      push edx

' *** Reference to "__vbaVarCat"
'007092f5    8b3508124000            mov esi, dword ptr [00401208]
'007092fb    ffd6                    call esi
'007092fd    50                      push eax
'007092fe    8d856cfeffff            lea eax, dword ptr [ebp+fffffe6c]
'00709304    50                      push eax
'00709305    8d8d1cffffff            lea ecx, dword ptr [ebp+ffffff1c]
'0070930b    51                      push ecx
'0070930c    ffd6                    call esi
'0070930e    50                      push eax
'0070930f    8d955cfeffff            lea edx, dword ptr [ebp+fffffe5c]
'00709315    52                      push edx
'00709316    8d850cffffff            lea eax, dword ptr [ebp+ffffff0c]
'0070931c    50                      push eax
'0070931d    ffd6                    call esi
'0070931f    50                      push eax
'00709320    8d8dfcfeffff            lea ecx, dword ptr [ebp+fffffefc]
'00709326    51                      push ecx
'00709327    8d95ecfeffff            lea edx, dword ptr [ebp+fffffeec]
'0070932d    52                      push edx
'0070932e    ffd6                    call esi
'00709330    50                      push eax
'00709331    8d854cfeffff            lea eax, dword ptr [ebp+fffffe4c]
'00709337    50                      push eax
'00709338    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'0070933e    51                      push ecx
'0070933f    ffd6                    call esi
'00709341    50                      push eax
'00709342    33d2                    xor edx, edx
'00709344    668b152ab07200          mov dx, word ptr [0072b02a]
'0070934b    52                      push edx
'0070934c    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'00709351    ff15b8114000            call dword ptr [004011b8]
Print #0, 
'00709357    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'0070935d    50                      push eax
'0070935e    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'00709364    51                      push ecx
'00709365    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00709367    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_91)
'0070936d    8d95dcfeffff            lea edx, dword ptr [ebp+fffffedc]
'00709373    52                      push edx
'00709374    8d85ecfeffff            lea eax, dword ptr [ebp+fffffeec]
'0070937a    50                      push eax
'0070937b    8d8dfcfeffff            lea ecx, dword ptr [ebp+fffffefc]
'00709381    51                      push ecx
'00709382    8d950cffffff            lea edx, dword ptr [ebp+ffffff0c]
'00709388    52                      push edx
'00709389    8d851cffffff            lea eax, dword ptr [ebp+ffffff1c]
'0070938f    50                      push eax
'00709390    8d8d2cffffff            lea ecx, dword ptr [ebp+ffffff2c]
'00709396    51                      push ecx
'00709397    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'0070939d    52                      push edx
'0070939e    8d854cffffff            lea eax, dword ptr [ebp+ffffff4c]
'007093a4    50                      push eax
'007093a5    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]
'007093ab    51                      push ecx
'007093ac    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'007093ae    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_88, var_89, var_90, var_93, var_95, var_116, var_117, var_118, var_10)
'007093b4    83c440                  add esp, 40
'007093b7    6a00                    push 00

' *** Reference to "__vbaResume"
'007093b9    ff1568104000            call dword ptr [00401068]
'007093bf    e986260000              jmp 70ba4a
Resume handler_70BA4A

' *** Reference to "__vbaHresultCheckObj"
'007093c4    8b1d80104000            mov ebx, dword ptr [00401080]
'007093ca    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'007093d0    33f6                    xor esi, esi
var_num8 = Empty
'007093d2    89b570ffffff            mov dword ptr [ebp+ffffff70], esi
'007093d8    50                      push eax
'007093d9    8d45d8                  lea eax, dword ptr [ebp-28]
'007093dc    50                      push eax

' *** Reference to "__vbaObjSet"
'007093dd    ff15b4104000            call dword ptr [004010b4]
Set var_45 = Nothing
'007093e3    8d4d84                  lea ecx, dword ptr [ebp-7c]
'007093e6    51                      push ecx
'007093e7    8d5588                  lea edx, dword ptr [ebp-78]
'007093ea    52                      push edx
'007093eb    8d458c                  lea eax, dword ptr [ebp-74]
'007093ee    50                      push eax
'007093ef    8d4d90                  lea ecx, dword ptr [ebp-70]
'007093f2    51                      push ecx
'007093f3    8d5594                  lea edx, dword ptr [ebp-6c]
'007093f6    52                      push edx
'007093f7    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'007093f9    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_121, -4536, -4540, -4544, -4548)
'007093ff    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'00709405    50                      push eax
'00709406    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0070940c    51                      push ecx
'0070940d    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070940f    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_91)
'00709415    83c424                  add esp, 24
'00709418    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]

' *** Reference to "__vbaFreeVar"
'0070941e    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_88)
'00709424    8b45a4                  mov eax, dword ptr [ebp-5c]
'00709427    8b10                    mov edx, dword ptr [eax]
'00709429    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0070942f    51                      push ecx
'00709430    50                      push eax
'00709431    ff92b4000000            call dword ptr [edx+000000b4]
'00709437    dbe2                    fnclex
'00709439    3bc6                    cmp eax, esi
'0070943b    7d11                    jge 70944e

If (var_83 < __vbaVarCat) Then
'0070943d    68b4000000              push 000000b4
'00709442    6830314300              push 00433130
'00709447    8b55a4                  mov edx, dword ptr [ebp-5c]
'0070944a    52                      push edx
'0070944b    50                      push eax
'0070944c    ffd3                    call ebx
    
End If
'0070944e    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'00709454    89858cfdffff            mov dword ptr [ebp+fffffd8c], eax
'0070945a    c78594feffff34d44300    mov dword ptr [ebp+fffffe94], 0043d434
'00709464    b908000000              mov ecx, 00000008
'00709469    898d8cfeffff            mov dword ptr [ebp+fffffe8c], ecx
'0070946f    8b10                    mov edx, dword ptr [eax]
'00709471    8dbd74ffffff            lea edi, dword ptr [ebp+ffffff74]
'00709477    57                      push edi
'00709478    83ec10                  sub esp, 10
'0070947b    8bfc                    mov edi, esp
'0070947d    890f                    mov dword ptr [edi], ecx
'0070947f    8b8d90feffff            mov ecx, dword ptr [ebp+fffffe90]
'00709485    894f04                  mov dword ptr [edi+04], ecx
'00709488    8b8d94feffff            mov ecx, dword ptr [ebp+fffffe94]
'0070948e    894f08                  mov dword ptr [edi+08], ecx
'00709491    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'00709497    894f0c                  mov dword ptr [edi+0c], ecx
'0070949a    50                      push eax
'0070949b    ff5230                  call dword ptr [edx+30]
var_87.Description = 
'0070949e    dbe2                    fnclex
'007094a0    3bc6                    cmp eax, esi
'007094a2    7d11                    jge 7094b5
'007094a4    6a30                    push 30
'007094a6    68d8304300              push 004330d8
'007094ab    8b958cfdffff            mov edx, dword ptr [ebp+fffffd8c]
'007094b1    52                      push edx
'007094b2    50                      push eax
'007094b3    ffd3                    call ebx
'007094b5    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'007094bb    89b574ffffff            mov dword ptr [ebp+ffffff74], esi
'007094c1    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'007094c7    c7855cffffff09000000    mov dword ptr [ebp+ffffff5c], 00000009
'007094d1    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'007094d7    50                      push eax

' *** Reference to "rtcIsNull"
'007094d8    ff1540114000            call dword ptr [00401140]
'007094de    898598fdffff            mov dword ptr [ebp+fffffd98], eax
'007094e4    8b45a4                  mov eax, dword ptr [ebp-5c]
'007094e7    8b08                    mov ecx, dword ptr [eax]
'007094e9    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'007094ef    52                      push edx
'007094f0    50                      push eax
'007094f1    ff91b4000000            call dword ptr [ecx+000000b4]
'007094f7    dbe2                    fnclex
'007094f9    3bc6                    cmp eax, esi
'007094fb    7d11                    jge 70950e

If (var_83 < __vbaVarCat) Then
'007094fd    68b4000000              push 000000b4
'00709502    6830314300              push 00433130
'00709507    8b4da4                  mov ecx, dword ptr [ebp-5c]
'0070950a    51                      push ecx
'0070950b    50                      push eax
'0070950c    ffd3                    call ebx
    
End If
'0070950e    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'00709514    898580fdffff            mov dword ptr [ebp+fffffd80], eax
'0070951a    c78584feffff34d44300    mov dword ptr [ebp+fffffe84], 0043d434
'00709524    b908000000              mov ecx, 00000008
'00709529    898d7cfeffff            mov dword ptr [ebp+fffffe7c], ecx
'0070952f    8b10                    mov edx, dword ptr [eax]
'00709531    8dbd6cffffff            lea edi, dword ptr [ebp+ffffff6c]
'00709537    57                      push edi
'00709538    83ec10                  sub esp, 10
'0070953b    8bfc                    mov edi, esp
'0070953d    890f                    mov dword ptr [edi], ecx
'0070953f    8b8d80feffff            mov ecx, dword ptr [ebp+fffffe80]
'00709545    894f04                  mov dword ptr [edi+04], ecx
'00709548    8b8d84feffff            mov ecx, dword ptr [ebp+fffffe84]
'0070954e    894f08                  mov dword ptr [edi+08], ecx
'00709551    8b8d88feffff            mov ecx, dword ptr [ebp+fffffe88]
'00709557    894f0c                  mov dword ptr [edi+0c], ecx
'0070955a    50                      push eax
'0070955b    ff5230                  call dword ptr [edx+30]
'0070955e    dbe2                    fnclex
'00709560    3bc6                    cmp eax, esi
'00709562    7d11                    jge 709575

If (__vbaVarCat < __vbaVarCat) Then
'00709564    6a30                    push 30
'00709566    68d8304300              push 004330d8
'0070956b    8b9580fdffff            mov edx, dword ptr [ebp+fffffd80]
'00709571    52                      push edx
'00709572    50                      push eax
'00709573    ffd3                    call ebx
    
End If
'00709575    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'0070957b    89b56cffffff            mov dword ptr [ebp+ffffff6c], esi
'00709581    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'00709587    c7853cffffff09000000    mov dword ptr [ebp+ffffff3c], 00000009
'00709591    89b554ffffff            mov dword ptr [ebp+ffffff54], esi
'00709597    c7854cffffff02000000    mov dword ptr [ebp+ffffff4c], 00000002
'007095a1    668b8598fdffff          mov ax, word ptr [ebp+fffffd98]
'007095a8    66898574feffff          mov word ptr [ebp+fffffe74], ax
'007095af    c7856cfeffff0b000000    mov dword ptr [ebp+fffffe6c], 0000000b
'007095b9    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'007095bf    51                      push ecx
'007095c0    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'007095c6    52                      push edx
'007095c7    8d856cfeffff            lea eax, dword ptr [ebp+fffffe6c]
'007095cd    50                      push eax
'007095ce    8d8d2cffffff            lea ecx, dword ptr [ebp+ffffff2c]
'007095d4    51                      push ecx

' *** Reference to "rtcImmediateIf"
'007095d5    ff154c124000            call dword ptr [0040124c]
'007095db    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'007095e1    52                      push edx

' *** Reference to "__vbaI2Var"
'007095e2    ff150c124000            call dword ptr [0040120c]
'007095e8    89459c                  mov dword ptr [ebp-64], eax
'007095eb    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'007095f1    50                      push eax
'007095f2    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'007095f8    51                      push ecx
'007095f9    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'007095fb    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_19)
'00709601    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'00709607    52                      push edx
'00709608    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'0070960e    50                      push eax
'0070960f    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'00709615    51                      push ecx
'00709616    8d956cfeffff            lea edx, dword ptr [ebp+fffffe6c]
'0070961c    52                      push edx
'0070961d    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'00709623    50                      push eax
'00709624    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'00709626    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_88, var_346, var_89, var_90, var_93)
'0070962c    83c424                  add esp, 24
'0070962f    8b45a4                  mov eax, dword ptr [ebp-5c]
'00709632    8b08                    mov ecx, dword ptr [eax]
'00709634    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0070963a    52                      push edx
'0070963b    50                      push eax
'0070963c    ff91b4000000            call dword ptr [ecx+000000b4]
'00709642    dbe2                    fnclex
'00709644    3bc6                    cmp eax, esi
'00709646    7d11                    jge 709659

If (var_83 < __vbaVarCat) Then
'00709648    68b4000000              push 000000b4
'0070964d    6830314300              push 00433130
'00709652    8b4da4                  mov ecx, dword ptr [ebp-5c]
'00709655    51                      push ecx
'00709656    50                      push eax
'00709657    ffd3                    call ebx
    
End If
'00709659    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'0070965f    89858cfdffff            mov dword ptr [ebp+fffffd8c], eax
'00709665    c78594feffff5cd44300    mov dword ptr [ebp+fffffe94], 0043d45c
'0070966f    b908000000              mov ecx, 00000008
'00709674    898d8cfeffff            mov dword ptr [ebp+fffffe8c], ecx
'0070967a    8b10                    mov edx, dword ptr [eax]
'0070967c    8dbd74ffffff            lea edi, dword ptr [ebp+ffffff74]
'00709682    57                      push edi
'00709683    83ec10                  sub esp, 10
'00709686    8bfc                    mov edi, esp
'00709688    890f                    mov dword ptr [edi], ecx
'0070968a    8b8d90feffff            mov ecx, dword ptr [ebp+fffffe90]
'00709690    894f04                  mov dword ptr [edi+04], ecx
'00709693    8b8d94feffff            mov ecx, dword ptr [ebp+fffffe94]
'00709699    894f08                  mov dword ptr [edi+08], ecx
'0070969c    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'007096a2    894f0c                  mov dword ptr [edi+0c], ecx
'007096a5    50                      push eax
'007096a6    ff5230                  call dword ptr [edx+30]
var_87.Description = 
'007096a9    dbe2                    fnclex
'007096ab    3bc6                    cmp eax, esi
'007096ad    7d11                    jge 7096c0
'007096af    6a30                    push 30
'007096b1    68d8304300              push 004330d8
'007096b6    8b958cfdffff            mov edx, dword ptr [ebp+fffffd8c]
'007096bc    52                      push edx
'007096bd    50                      push eax
'007096be    ffd3                    call ebx
'007096c0    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'007096c6    89b574ffffff            mov dword ptr [ebp+ffffff74], esi
'007096cc    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'007096d2    c7855cffffff09000000    mov dword ptr [ebp+ffffff5c], 00000009
'007096dc    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'007096e2    50                      push eax

' *** Reference to "rtcIsNull"
'007096e3    ff1540114000            call dword ptr [00401140]
'007096e9    898598fdffff            mov dword ptr [ebp+fffffd98], eax
'007096ef    8b45a4                  mov eax, dword ptr [ebp-5c]
'007096f2    8b08                    mov ecx, dword ptr [eax]
'007096f4    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'007096fa    52                      push edx
'007096fb    50                      push eax
'007096fc    ff91b4000000            call dword ptr [ecx+000000b4]
'00709702    dbe2                    fnclex
'00709704    3bc6                    cmp eax, esi
'00709706    7d11                    jge 709719

If (var_83 < __vbaVarCat) Then
'00709708    68b4000000              push 000000b4
'0070970d    6830314300              push 00433130
'00709712    8b4da4                  mov ecx, dword ptr [ebp-5c]
'00709715    51                      push ecx
'00709716    50                      push eax
'00709717    ffd3                    call ebx
    
End If
'00709719    8b8570ffffff            mov eax, dword ptr [ebp+ffffff70]
'0070971f    898580fdffff            mov dword ptr [ebp+fffffd80], eax
'00709725    c78584feffff5cd44300    mov dword ptr [ebp+fffffe84], 0043d45c
'0070972f    b908000000              mov ecx, 00000008
'00709734    898d7cfeffff            mov dword ptr [ebp+fffffe7c], ecx
'0070973a    8b10                    mov edx, dword ptr [eax]
'0070973c    8dbd6cffffff            lea edi, dword ptr [ebp+ffffff6c]
'00709742    57                      push edi
'00709743    83ec10                  sub esp, 10
'00709746    8bfc                    mov edi, esp
'00709748    890f                    mov dword ptr [edi], ecx
'0070974a    8b8d80feffff            mov ecx, dword ptr [ebp+fffffe80]
'00709750    894f04                  mov dword ptr [edi+04], ecx
'00709753    8b8d84feffff            mov ecx, dword ptr [ebp+fffffe84]
'00709759    894f08                  mov dword ptr [edi+08], ecx
'0070975c    8b8d88feffff            mov ecx, dword ptr [ebp+fffffe88]
'00709762    894f0c                  mov dword ptr [edi+0c], ecx
'00709765    50                      push eax
'00709766    ff5230                  call dword ptr [edx+30]
'00709769    dbe2                    fnclex
'0070976b    3bc6                    cmp eax, esi
'0070976d    7d11                    jge 709780

If (__vbaVarCat < __vbaVarCat) Then
'0070976f    6a30                    push 30
'00709771    68d8304300              push 004330d8
'00709776    8b9580fdffff            mov edx, dword ptr [ebp+fffffd80]
'0070977c    52                      push edx
'0070977d    50                      push eax
'0070977e    ffd3                    call ebx
    
End If
'00709780    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'00709786    89b56cffffff            mov dword ptr [ebp+ffffff6c], esi
'0070978c    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'00709792    c7853cffffff09000000    mov dword ptr [ebp+ffffff3c], 00000009
'0070979c    89b554ffffff            mov dword ptr [ebp+ffffff54], esi
'007097a2    c7854cffffff02000000    mov dword ptr [ebp+ffffff4c], 00000002
'007097ac    668b8598fdffff          mov ax, word ptr [ebp+fffffd98]
'007097b3    66898574feffff          mov word ptr [ebp+fffffe74], ax
'007097ba    c7856cfeffff0b000000    mov dword ptr [ebp+fffffe6c], 0000000b
'007097c4    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'007097ca    51                      push ecx
'007097cb    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'007097d1    52                      push edx
'007097d2    8d856cfeffff            lea eax, dword ptr [ebp+fffffe6c]
'007097d8    50                      push eax
'007097d9    8d8d2cffffff            lea ecx, dword ptr [ebp+ffffff2c]
'007097df    51                      push ecx

' *** Reference to "rtcImmediateIf"
'007097e0    ff154c124000            call dword ptr [0040124c]
'007097e6    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'007097ec    52                      push edx

' *** Reference to "__vbaI2Var"
'007097ed    ff150c124000            call dword ptr [0040120c]
'007097f3    8bf8                    mov edi, eax
'007097f5    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'007097fb    50                      push eax
'007097fc    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'00709802    51                      push ecx
'00709803    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00709805    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_19)
'0070980b    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'00709811    52                      push edx
'00709812    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'00709818    50                      push eax
'00709819    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'0070981f    51                      push ecx
'00709820    8d956cfeffff            lea edx, dword ptr [ebp+fffffe6c]
'00709826    52                      push edx
'00709827    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070982d    50                      push eax
'0070982e    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'00709830    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_88, var_346, var_89, var_90, var_93)
'00709836    83c424                  add esp, 24
'00709839    8b45a4                  mov eax, dword ptr [ebp-5c]
'0070983c    8b08                    mov ecx, dword ptr [eax]
'0070983e    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'00709844    52                      push edx
'00709845    50                      push eax
'00709846    ff91b4000000            call dword ptr [ecx+000000b4]
'0070984c    dbe2                    fnclex
'0070984e    3bc6                    cmp eax, esi
'00709850    7d11                    jge 709863

If (var_83 < __vbaVarCat) Then
'00709852    68b4000000              push 000000b4
'00709857    6830314300              push 00433130
'0070985c    8b4da4                  mov ecx, dword ptr [ebp-5c]
'0070985f    51                      push ecx
'00709860    50                      push eax
'00709861    ffd3                    call ebx
    
End If
'00709863    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'00709869    898d8cfdffff            mov dword ptr [ebp+fffffd8c], ecx
'0070986f    c78594feffff84d44300    mov dword ptr [ebp+fffffe94], 0043d484
'00709879    b808000000              mov eax, 00000008
'0070987e    89858cfeffff            mov dword ptr [ebp+fffffe8c], eax
'00709884    8b11                    mov edx, dword ptr [ecx]
'00709886    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'0070988c    51                      push ecx
'0070988d    83ec10                  sub esp, 10
'00709890    8bcc                    mov ecx, esp
'00709892    8901                    mov dword ptr [ecx], eax
'00709894    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'0070989a    894104                  mov dword ptr [ecx+04], eax
'0070989d    8b8594feffff            mov eax, dword ptr [ebp+fffffe94]
'007098a3    894108                  mov dword ptr [ecx+08], eax
'007098a6    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'007098ac    89410c                  mov dword ptr [ecx+0c], eax
'007098af    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'007098b5    51                      push ecx
'007098b6    ff5230                  call dword ptr [edx+30]
var_87.Description = 
'007098b9    dbe2                    fnclex
'007098bb    3bc6                    cmp eax, esi
'007098bd    7d11                    jge 7098d0
'007098bf    6a30                    push 30
'007098c1    68d8304300              push 004330d8
'007098c6    8b958cfdffff            mov edx, dword ptr [ebp+fffffd8c]
'007098cc    52                      push edx
'007098cd    50                      push eax
'007098ce    ffd3                    call ebx
'007098d0    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'007098d6    89b574ffffff            mov dword ptr [ebp+ffffff74], esi
'007098dc    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'007098e2    c7855cffffff09000000    mov dword ptr [ebp+ffffff5c], 00000009
'007098ec    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'007098f2    50                      push eax

' *** Reference to "rtcIsNull"
'007098f3    ff1540114000            call dword ptr [00401140]
'007098f9    898598fdffff            mov dword ptr [ebp+fffffd98], eax
'007098ff    8b45a4                  mov eax, dword ptr [ebp-5c]
'00709902    8b08                    mov ecx, dword ptr [eax]
'00709904    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0070990a    52                      push edx
'0070990b    50                      push eax
'0070990c    ff91b4000000            call dword ptr [ecx+000000b4]
'00709912    dbe2                    fnclex
'00709914    3bc6                    cmp eax, esi
'00709916    7d11                    jge 709929

If (var_83 < __vbaVarCat) Then
'00709918    68b4000000              push 000000b4
'0070991d    6830314300              push 00433130
'00709922    8b4da4                  mov ecx, dword ptr [ebp-5c]
'00709925    51                      push ecx
'00709926    50                      push eax
'00709927    ffd3                    call ebx
    
End If
'00709929    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'0070992f    898d80fdffff            mov dword ptr [ebp+fffffd80], ecx
'00709935    c78584feffff84d44300    mov dword ptr [ebp+fffffe84], 0043d484
'0070993f    b808000000              mov eax, 00000008
'00709944    89857cfeffff            mov dword ptr [ebp+fffffe7c], eax
'0070994a    8b11                    mov edx, dword ptr [ecx]
'0070994c    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'00709952    51                      push ecx
'00709953    83ec10                  sub esp, 10
'00709956    8bcc                    mov ecx, esp
'00709958    8901                    mov dword ptr [ecx], eax
'0070995a    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'00709960    894104                  mov dword ptr [ecx+04], eax
'00709963    8b8584feffff            mov eax, dword ptr [ebp+fffffe84]
'00709969    894108                  mov dword ptr [ecx+08], eax
'0070996c    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'00709972    89410c                  mov dword ptr [ecx+0c], eax
'00709975    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'0070997b    51                      push ecx
'0070997c    ff5230                  call dword ptr [edx+30]
'0070997f    dbe2                    fnclex
'00709981    3bc6                    cmp eax, esi
'00709983    7d11                    jge 709996

If (0 < __vbaVarCat) Then
'00709985    6a30                    push 30
'00709987    68d8304300              push 004330d8
'0070998c    8b9580fdffff            mov edx, dword ptr [ebp+fffffd80]
'00709992    52                      push edx
'00709993    50                      push eax
'00709994    ffd3                    call ebx
    
End If
'00709996    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'0070999c    89b56cffffff            mov dword ptr [ebp+ffffff6c], esi
'007099a2    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'007099a8    c7853cffffff09000000    mov dword ptr [ebp+ffffff3c], 00000009
'007099b2    89b554ffffff            mov dword ptr [ebp+ffffff54], esi
'007099b8    c7854cffffff02000000    mov dword ptr [ebp+ffffff4c], 00000002
'007099c2    668b8598fdffff          mov ax, word ptr [ebp+fffffd98]
'007099c9    66898574feffff          mov word ptr [ebp+fffffe74], ax
'007099d0    c7856cfeffff0b000000    mov dword ptr [ebp+fffffe6c], 0000000b
'007099da    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'007099e0    51                      push ecx
'007099e1    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'007099e7    52                      push edx
'007099e8    8d856cfeffff            lea eax, dword ptr [ebp+fffffe6c]
'007099ee    50                      push eax
'007099ef    8d8d2cffffff            lea ecx, dword ptr [ebp+ffffff2c]
'007099f5    51                      push ecx

' *** Reference to "rtcImmediateIf"
'007099f6    ff154c124000            call dword ptr [0040124c]
'007099fc    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'00709a02    52                      push edx

' *** Reference to "__vbaI2Var"
'00709a03    ff150c124000            call dword ptr [0040120c]
'00709a09    8945dc                  mov dword ptr [ebp-24], eax
'00709a0c    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'00709a12    50                      push eax
'00709a13    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'00709a19    51                      push ecx
'00709a1a    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00709a1c    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_19)
'00709a22    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'00709a28    52                      push edx
'00709a29    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'00709a2f    50                      push eax
'00709a30    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'00709a36    51                      push ecx
'00709a37    8d956cfeffff            lea edx, dword ptr [ebp+fffffe6c]
'00709a3d    52                      push edx
'00709a3e    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'00709a44    50                      push eax
'00709a45    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'00709a47    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_88, var_346, var_89, var_90, var_93)
'00709a4d    83c424                  add esp, 24
'00709a50    8b45a4                  mov eax, dword ptr [ebp-5c]
'00709a53    8b08                    mov ecx, dword ptr [eax]
'00709a55    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'00709a5b    52                      push edx
'00709a5c    50                      push eax
'00709a5d    ff91b4000000            call dword ptr [ecx+000000b4]
'00709a63    dbe2                    fnclex
'00709a65    3bc6                    cmp eax, esi
'00709a67    7d11                    jge 709a7a

If (var_83 < __vbaVarCat) Then
'00709a69    68b4000000              push 000000b4
'00709a6e    6830314300              push 00433130
'00709a73    8b4da4                  mov ecx, dword ptr [ebp-5c]
'00709a76    51                      push ecx
'00709a77    50                      push eax
'00709a78    ffd3                    call ebx
    
End If
'00709a7a    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'00709a80    898d8cfdffff            mov dword ptr [ebp+fffffd8c], ecx
'00709a86    c78594feffff08ab4300    mov dword ptr [ebp+fffffe94], 0043ab08
'00709a90    b808000000              mov eax, 00000008
'00709a95    89858cfeffff            mov dword ptr [ebp+fffffe8c], eax
'00709a9b    8b11                    mov edx, dword ptr [ecx]
'00709a9d    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'00709aa3    51                      push ecx
'00709aa4    83ec10                  sub esp, 10
'00709aa7    8bcc                    mov ecx, esp
'00709aa9    8901                    mov dword ptr [ecx], eax
'00709aab    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'00709ab1    894104                  mov dword ptr [ecx+04], eax
'00709ab4    8b8594feffff            mov eax, dword ptr [ebp+fffffe94]
'00709aba    894108                  mov dword ptr [ecx+08], eax
'00709abd    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'00709ac3    89410c                  mov dword ptr [ecx+0c], eax
'00709ac6    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'00709acc    51                      push ecx
'00709acd    ff5230                  call dword ptr [edx+30]
var_87.Description = 
'00709ad0    dbe2                    fnclex
'00709ad2    3bc6                    cmp eax, esi
'00709ad4    7d11                    jge 709ae7
'00709ad6    6a30                    push 30
'00709ad8    68d8304300              push 004330d8
'00709add    8b958cfdffff            mov edx, dword ptr [ebp+fffffd8c]
'00709ae3    52                      push edx
'00709ae4    50                      push eax
'00709ae5    ffd3                    call ebx
'00709ae7    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'00709aed    89b574ffffff            mov dword ptr [ebp+ffffff74], esi
'00709af3    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'00709af9    c7855cffffff09000000    mov dword ptr [ebp+ffffff5c], 00000009
'00709b03    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'00709b09    50                      push eax

' *** Reference to "rtcIsNull"
'00709b0a    ff1540114000            call dword ptr [00401140]
'00709b10    898598fdffff            mov dword ptr [ebp+fffffd98], eax
'00709b16    8b45a4                  mov eax, dword ptr [ebp-5c]
'00709b19    8b08                    mov ecx, dword ptr [eax]
'00709b1b    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'00709b21    52                      push edx
'00709b22    50                      push eax
'00709b23    ff91b4000000            call dword ptr [ecx+000000b4]
'00709b29    dbe2                    fnclex
'00709b2b    3bc6                    cmp eax, esi
'00709b2d    7d11                    jge 709b40

If (var_83 < __vbaVarCat) Then
'00709b2f    68b4000000              push 000000b4
'00709b34    6830314300              push 00433130
'00709b39    8b4da4                  mov ecx, dword ptr [ebp-5c]
'00709b3c    51                      push ecx
'00709b3d    50                      push eax
'00709b3e    ffd3                    call ebx
    
End If
'00709b40    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'00709b46    898d80fdffff            mov dword ptr [ebp+fffffd80], ecx
'00709b4c    c78584feffff08ab4300    mov dword ptr [ebp+fffffe84], 0043ab08
'00709b56    b808000000              mov eax, 00000008
'00709b5b    89857cfeffff            mov dword ptr [ebp+fffffe7c], eax
'00709b61    8b11                    mov edx, dword ptr [ecx]
'00709b63    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'00709b69    51                      push ecx
'00709b6a    83ec10                  sub esp, 10
'00709b6d    8bcc                    mov ecx, esp
'00709b6f    8901                    mov dword ptr [ecx], eax
'00709b71    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'00709b77    894104                  mov dword ptr [ecx+04], eax
'00709b7a    8b8584feffff            mov eax, dword ptr [ebp+fffffe84]
'00709b80    894108                  mov dword ptr [ecx+08], eax
'00709b83    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'00709b89    89410c                  mov dword ptr [ecx+0c], eax
'00709b8c    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'00709b92    51                      push ecx
'00709b93    ff5230                  call dword ptr [edx+30]
'00709b96    dbe2                    fnclex
'00709b98    3bc6                    cmp eax, esi
'00709b9a    7d11                    jge 709bad

If (0 < __vbaVarCat) Then
'00709b9c    6a30                    push 30
'00709b9e    68d8304300              push 004330d8
'00709ba3    8b9580fdffff            mov edx, dword ptr [ebp+fffffd80]
'00709ba9    52                      push edx
'00709baa    50                      push eax
'00709bab    ffd3                    call ebx
    
End If
'00709bad    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'00709bb3    89b56cffffff            mov dword ptr [ebp+ffffff6c], esi
'00709bb9    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'00709bbf    c7853cffffff09000000    mov dword ptr [ebp+ffffff3c], 00000009
'00709bc9    89b554ffffff            mov dword ptr [ebp+ffffff54], esi
'00709bcf    c7854cffffff02000000    mov dword ptr [ebp+ffffff4c], 00000002
'00709bd9    668b8598fdffff          mov ax, word ptr [ebp+fffffd98]
'00709be0    66898574feffff          mov word ptr [ebp+fffffe74], ax
'00709be7    c7856cfeffff0b000000    mov dword ptr [ebp+fffffe6c], 0000000b
'00709bf1    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'00709bf7    51                      push ecx
'00709bf8    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'00709bfe    52                      push edx
'00709bff    8d856cfeffff            lea eax, dword ptr [ebp+fffffe6c]
'00709c05    50                      push eax
'00709c06    8d8d2cffffff            lea ecx, dword ptr [ebp+ffffff2c]
'00709c0c    51                      push ecx

' *** Reference to "rtcImmediateIf"
'00709c0d    ff154c124000            call dword ptr [0040124c]
'00709c13    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'00709c19    52                      push edx

' *** Reference to "__vbaI2Var"
'00709c1a    ff150c124000            call dword ptr [0040120c]
'00709c20    8945d4                  mov dword ptr [ebp-2c], eax
'00709c23    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'00709c29    50                      push eax
'00709c2a    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'00709c30    51                      push ecx
'00709c31    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00709c33    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_19)
'00709c39    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'00709c3f    52                      push edx
'00709c40    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'00709c46    50                      push eax
'00709c47    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'00709c4d    51                      push ecx
'00709c4e    8d956cfeffff            lea edx, dword ptr [ebp+fffffe6c]
'00709c54    52                      push edx
'00709c55    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'00709c5b    50                      push eax
'00709c5c    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'00709c5e    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_88, var_346, var_89, var_90, var_93)
'00709c64    83c424                  add esp, 24
'00709c67    8b45a4                  mov eax, dword ptr [ebp-5c]
'00709c6a    8b08                    mov ecx, dword ptr [eax]
'00709c6c    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'00709c72    52                      push edx
'00709c73    50                      push eax
'00709c74    ff91b4000000            call dword ptr [ecx+000000b4]
'00709c7a    dbe2                    fnclex
'00709c7c    3bc6                    cmp eax, esi
'00709c7e    7d11                    jge 709c91

If (var_83 < __vbaVarCat) Then
'00709c80    68b4000000              push 000000b4
'00709c85    6830314300              push 00433130
'00709c8a    8b4da4                  mov ecx, dword ptr [ebp-5c]
'00709c8d    51                      push ecx
'00709c8e    50                      push eax
'00709c8f    ffd3                    call ebx
    
End If
'00709c91    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'00709c97    898d8cfdffff            mov dword ptr [ebp+fffffd8c], ecx
'00709c9d    c78594feffff18ab4300    mov dword ptr [ebp+fffffe94], 0043ab18
'00709ca7    b808000000              mov eax, 00000008
'00709cac    89858cfeffff            mov dword ptr [ebp+fffffe8c], eax
'00709cb2    8b11                    mov edx, dword ptr [ecx]
'00709cb4    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'00709cba    51                      push ecx
'00709cbb    83ec10                  sub esp, 10
'00709cbe    8bcc                    mov ecx, esp
'00709cc0    8901                    mov dword ptr [ecx], eax
'00709cc2    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'00709cc8    894104                  mov dword ptr [ecx+04], eax
'00709ccb    8b8594feffff            mov eax, dword ptr [ebp+fffffe94]
'00709cd1    894108                  mov dword ptr [ecx+08], eax
'00709cd4    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'00709cda    89410c                  mov dword ptr [ecx+0c], eax
'00709cdd    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'00709ce3    51                      push ecx
'00709ce4    ff5230                  call dword ptr [edx+30]
var_87.Description = 
'00709ce7    dbe2                    fnclex
'00709ce9    3bc6                    cmp eax, esi
'00709ceb    7d11                    jge 709cfe
'00709ced    6a30                    push 30
'00709cef    68d8304300              push 004330d8
'00709cf4    8b958cfdffff            mov edx, dword ptr [ebp+fffffd8c]
'00709cfa    52                      push edx
'00709cfb    50                      push eax
'00709cfc    ffd3                    call ebx
'00709cfe    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'00709d04    89b574ffffff            mov dword ptr [ebp+ffffff74], esi
'00709d0a    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'00709d10    c7855cffffff09000000    mov dword ptr [ebp+ffffff5c], 00000009
'00709d1a    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'00709d20    50                      push eax

' *** Reference to "rtcIsNull"
'00709d21    ff1540114000            call dword ptr [00401140]
'00709d27    898598fdffff            mov dword ptr [ebp+fffffd98], eax
'00709d2d    8b45a4                  mov eax, dword ptr [ebp-5c]
'00709d30    8b08                    mov ecx, dword ptr [eax]
'00709d32    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'00709d38    52                      push edx
'00709d39    50                      push eax
'00709d3a    ff91b4000000            call dword ptr [ecx+000000b4]
'00709d40    dbe2                    fnclex
'00709d42    3bc6                    cmp eax, esi
'00709d44    7d11                    jge 709d57

If (var_83 < __vbaVarCat) Then
'00709d46    68b4000000              push 000000b4
'00709d4b    6830314300              push 00433130
'00709d50    8b4da4                  mov ecx, dword ptr [ebp-5c]
'00709d53    51                      push ecx
'00709d54    50                      push eax
'00709d55    ffd3                    call ebx
    
End If
'00709d57    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'00709d5d    898d80fdffff            mov dword ptr [ebp+fffffd80], ecx
'00709d63    c78584feffff18ab4300    mov dword ptr [ebp+fffffe84], 0043ab18
'00709d6d    b808000000              mov eax, 00000008
'00709d72    89857cfeffff            mov dword ptr [ebp+fffffe7c], eax
'00709d78    8b11                    mov edx, dword ptr [ecx]
'00709d7a    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'00709d80    51                      push ecx
'00709d81    83ec10                  sub esp, 10
'00709d84    8bcc                    mov ecx, esp
'00709d86    8901                    mov dword ptr [ecx], eax
'00709d88    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'00709d8e    894104                  mov dword ptr [ecx+04], eax
'00709d91    8b8584feffff            mov eax, dword ptr [ebp+fffffe84]
'00709d97    894108                  mov dword ptr [ecx+08], eax
'00709d9a    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'00709da0    89410c                  mov dword ptr [ecx+0c], eax
'00709da3    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'00709da9    51                      push ecx
'00709daa    ff5230                  call dword ptr [edx+30]
'00709dad    dbe2                    fnclex
'00709daf    3bc6                    cmp eax, esi
'00709db1    7d11                    jge 709dc4

If (0 < __vbaVarCat) Then
'00709db3    6a30                    push 30
'00709db5    68d8304300              push 004330d8
'00709dba    8b9580fdffff            mov edx, dword ptr [ebp+fffffd80]
'00709dc0    52                      push edx
'00709dc1    50                      push eax
'00709dc2    ffd3                    call ebx
    
End If
'00709dc4    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'00709dca    89b56cffffff            mov dword ptr [ebp+ffffff6c], esi
'00709dd0    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'00709dd6    c7853cffffff09000000    mov dword ptr [ebp+ffffff3c], 00000009
'00709de0    89b554ffffff            mov dword ptr [ebp+ffffff54], esi
'00709de6    c7854cffffff02000000    mov dword ptr [ebp+ffffff4c], 00000002
'00709df0    668b8598fdffff          mov ax, word ptr [ebp+fffffd98]
'00709df7    66898574feffff          mov word ptr [ebp+fffffe74], ax
'00709dfe    c7856cfeffff0b000000    mov dword ptr [ebp+fffffe6c], 0000000b
'00709e08    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'00709e0e    51                      push ecx
'00709e0f    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'00709e15    52                      push edx
'00709e16    8d856cfeffff            lea eax, dword ptr [ebp+fffffe6c]
'00709e1c    50                      push eax
'00709e1d    8d8d2cffffff            lea ecx, dword ptr [ebp+ffffff2c]
'00709e23    51                      push ecx

' *** Reference to "rtcImmediateIf"
'00709e24    ff154c124000            call dword ptr [0040124c]
'00709e2a    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'00709e30    52                      push edx

' *** Reference to "__vbaI2Var"
'00709e31    ff150c124000            call dword ptr [0040120c]
'00709e37    8945d0                  mov dword ptr [ebp-30], eax
'00709e3a    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'00709e40    50                      push eax
'00709e41    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'00709e47    51                      push ecx
'00709e48    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'00709e4a    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_19)
'00709e50    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'00709e56    52                      push edx
'00709e57    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'00709e5d    50                      push eax
'00709e5e    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'00709e64    51                      push ecx
'00709e65    8d956cfeffff            lea edx, dword ptr [ebp+fffffe6c]
'00709e6b    52                      push edx
'00709e6c    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'00709e72    50                      push eax
'00709e73    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'00709e75    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_88, var_346, var_89, var_90, var_93)
'00709e7b    83c424                  add esp, 24
'00709e7e    8b45a4                  mov eax, dword ptr [ebp-5c]
'00709e81    8b08                    mov ecx, dword ptr [eax]
'00709e83    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'00709e89    52                      push edx
'00709e8a    50                      push eax
'00709e8b    ff91b4000000            call dword ptr [ecx+000000b4]
'00709e91    dbe2                    fnclex
'00709e93    3bc6                    cmp eax, esi
'00709e95    7d11                    jge 709ea8

If (var_83 < __vbaVarCat) Then
'00709e97    68b4000000              push 000000b4
'00709e9c    6830314300              push 00433130
'00709ea1    8b4da4                  mov ecx, dword ptr [ebp-5c]
'00709ea4    51                      push ecx
'00709ea5    50                      push eax
'00709ea6    ffd3                    call ebx
    
End If
'00709ea8    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'00709eae    898d8cfdffff            mov dword ptr [ebp+fffffd8c], ecx
'00709eb4    c78594feffff28ab4300    mov dword ptr [ebp+fffffe94], 0043ab28
'00709ebe    b808000000              mov eax, 00000008
'00709ec3    89858cfeffff            mov dword ptr [ebp+fffffe8c], eax
'00709ec9    8b11                    mov edx, dword ptr [ecx]
'00709ecb    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'00709ed1    51                      push ecx
'00709ed2    83ec10                  sub esp, 10
'00709ed5    8bcc                    mov ecx, esp
'00709ed7    8901                    mov dword ptr [ecx], eax
'00709ed9    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'00709edf    894104                  mov dword ptr [ecx+04], eax
'00709ee2    8b8594feffff            mov eax, dword ptr [ebp+fffffe94]
'00709ee8    894108                  mov dword ptr [ecx+08], eax
'00709eeb    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'00709ef1    89410c                  mov dword ptr [ecx+0c], eax
'00709ef4    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'00709efa    51                      push ecx
'00709efb    ff5230                  call dword ptr [edx+30]
var_87.Description = 
'00709efe    dbe2                    fnclex
'00709f00    3bc6                    cmp eax, esi
'00709f02    7d11                    jge 709f15
'00709f04    6a30                    push 30
'00709f06    68d8304300              push 004330d8
'00709f0b    8b958cfdffff            mov edx, dword ptr [ebp+fffffd8c]
'00709f11    52                      push edx
'00709f12    50                      push eax
'00709f13    ffd3                    call ebx
'00709f15    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'00709f1b    89b574ffffff            mov dword ptr [ebp+ffffff74], esi
'00709f21    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'00709f27    c7855cffffff09000000    mov dword ptr [ebp+ffffff5c], 00000009
'00709f31    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'00709f37    50                      push eax

' *** Reference to "rtcIsNull"
'00709f38    ff1540114000            call dword ptr [00401140]
'00709f3e    898598fdffff            mov dword ptr [ebp+fffffd98], eax
'00709f44    8b45a4                  mov eax, dword ptr [ebp-5c]
'00709f47    8b08                    mov ecx, dword ptr [eax]
'00709f49    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'00709f4f    52                      push edx
'00709f50    50                      push eax
'00709f51    ff91b4000000            call dword ptr [ecx+000000b4]
'00709f57    dbe2                    fnclex
'00709f59    3bc6                    cmp eax, esi
'00709f5b    7d11                    jge 709f6e

If (var_83 < __vbaVarCat) Then
'00709f5d    68b4000000              push 000000b4
'00709f62    6830314300              push 00433130
'00709f67    8b4da4                  mov ecx, dword ptr [ebp-5c]
'00709f6a    51                      push ecx
'00709f6b    50                      push eax
'00709f6c    ffd3                    call ebx
    
End If
'00709f6e    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'00709f74    898d80fdffff            mov dword ptr [ebp+fffffd80], ecx
'00709f7a    c78584feffff28ab4300    mov dword ptr [ebp+fffffe84], 0043ab28
'00709f84    b808000000              mov eax, 00000008
'00709f89    89857cfeffff            mov dword ptr [ebp+fffffe7c], eax
'00709f8f    8b11                    mov edx, dword ptr [ecx]
'00709f91    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'00709f97    51                      push ecx
'00709f98    83ec10                  sub esp, 10
'00709f9b    8bcc                    mov ecx, esp
'00709f9d    8901                    mov dword ptr [ecx], eax
'00709f9f    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'00709fa5    894104                  mov dword ptr [ecx+04], eax
'00709fa8    8b8584feffff            mov eax, dword ptr [ebp+fffffe84]
'00709fae    894108                  mov dword ptr [ecx+08], eax
'00709fb1    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'00709fb7    89410c                  mov dword ptr [ecx+0c], eax
'00709fba    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'00709fc0    51                      push ecx
'00709fc1    ff5230                  call dword ptr [edx+30]
'00709fc4    dbe2                    fnclex
'00709fc6    3bc6                    cmp eax, esi
'00709fc8    7d11                    jge 709fdb

If (0 < __vbaVarCat) Then
'00709fca    6a30                    push 30
'00709fcc    68d8304300              push 004330d8
'00709fd1    8b9580fdffff            mov edx, dword ptr [ebp+fffffd80]
'00709fd7    52                      push edx
'00709fd8    50                      push eax
'00709fd9    ffd3                    call ebx
    
End If
'00709fdb    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'00709fe1    89b56cffffff            mov dword ptr [ebp+ffffff6c], esi
'00709fe7    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'00709fed    c7853cffffff09000000    mov dword ptr [ebp+ffffff3c], 00000009
'00709ff7    89b554ffffff            mov dword ptr [ebp+ffffff54], esi
'00709ffd    c7854cffffff02000000    mov dword ptr [ebp+ffffff4c], 00000002
'0070a007    668b8598fdffff          mov ax, word ptr [ebp+fffffd98]
'0070a00e    66898574feffff          mov word ptr [ebp+fffffe74], ax
'0070a015    c7856cfeffff0b000000    mov dword ptr [ebp+fffffe6c], 0000000b
'0070a01f    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'0070a025    51                      push ecx
'0070a026    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'0070a02c    52                      push edx
'0070a02d    8d856cfeffff            lea eax, dword ptr [ebp+fffffe6c]
'0070a033    50                      push eax
'0070a034    8d8d2cffffff            lea ecx, dword ptr [ebp+ffffff2c]
'0070a03a    51                      push ecx

' *** Reference to "rtcImmediateIf"
'0070a03b    ff154c124000            call dword ptr [0040124c]
'0070a041    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'0070a047    52                      push edx

' *** Reference to "__vbaI2Var"
'0070a048    ff150c124000            call dword ptr [0040120c]
'0070a04e    8945cc                  mov dword ptr [ebp-34], eax
'0070a051    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'0070a057    50                      push eax
'0070a058    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0070a05e    51                      push ecx
'0070a05f    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070a061    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_19)
'0070a067    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'0070a06d    52                      push edx
'0070a06e    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'0070a074    50                      push eax
'0070a075    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'0070a07b    51                      push ecx
'0070a07c    8d956cfeffff            lea edx, dword ptr [ebp+fffffe6c]
'0070a082    52                      push edx
'0070a083    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070a089    50                      push eax
'0070a08a    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0070a08c    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_88, var_346, var_89, var_90, var_93)
'0070a092    83c424                  add esp, 24
'0070a095    8b45a4                  mov eax, dword ptr [ebp-5c]
'0070a098    8b08                    mov ecx, dword ptr [eax]
'0070a09a    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0070a0a0    52                      push edx
'0070a0a1    50                      push eax
'0070a0a2    ff91b4000000            call dword ptr [ecx+000000b4]
'0070a0a8    dbe2                    fnclex
'0070a0aa    3bc6                    cmp eax, esi
'0070a0ac    7d11                    jge 70a0bf

If (var_83 < __vbaVarCat) Then
'0070a0ae    68b4000000              push 000000b4
'0070a0b3    6830314300              push 00433130
'0070a0b8    8b4da4                  mov ecx, dword ptr [ebp-5c]
'0070a0bb    51                      push ecx
'0070a0bc    50                      push eax
'0070a0bd    ffd3                    call ebx
    
End If
'0070a0bf    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'0070a0c5    898d8cfdffff            mov dword ptr [ebp+fffffd8c], ecx
'0070a0cb    c78594feffffac874300    mov dword ptr [ebp+fffffe94], 004387ac
'0070a0d5    b808000000              mov eax, 00000008
'0070a0da    89858cfeffff            mov dword ptr [ebp+fffffe8c], eax
'0070a0e0    8b11                    mov edx, dword ptr [ecx]
'0070a0e2    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'0070a0e8    51                      push ecx
'0070a0e9    83ec10                  sub esp, 10
'0070a0ec    8bcc                    mov ecx, esp
'0070a0ee    8901                    mov dword ptr [ecx], eax
'0070a0f0    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'0070a0f6    894104                  mov dword ptr [ecx+04], eax
'0070a0f9    8b8594feffff            mov eax, dword ptr [ebp+fffffe94]
'0070a0ff    894108                  mov dword ptr [ecx+08], eax
'0070a102    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0070a108    89410c                  mov dword ptr [ecx+0c], eax
'0070a10b    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'0070a111    51                      push ecx
'0070a112    ff5230                  call dword ptr [edx+30]
var_87.Description = 
'0070a115    dbe2                    fnclex
'0070a117    3bc6                    cmp eax, esi
'0070a119    7d11                    jge 70a12c
'0070a11b    6a30                    push 30
'0070a11d    68d8304300              push 004330d8
'0070a122    8b958cfdffff            mov edx, dword ptr [ebp+fffffd8c]
'0070a128    52                      push edx
'0070a129    50                      push eax
'0070a12a    ffd3                    call ebx
'0070a12c    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'0070a132    89b574ffffff            mov dword ptr [ebp+ffffff74], esi
'0070a138    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'0070a13e    c7855cffffff09000000    mov dword ptr [ebp+ffffff5c], 00000009
'0070a148    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070a14e    50                      push eax

' *** Reference to "rtcIsNull"
'0070a14f    ff1540114000            call dword ptr [00401140]
'0070a155    898598fdffff            mov dword ptr [ebp+fffffd98], eax
'0070a15b    8b45a4                  mov eax, dword ptr [ebp-5c]
'0070a15e    8b08                    mov ecx, dword ptr [eax]
'0070a160    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0070a166    52                      push edx
'0070a167    50                      push eax
'0070a168    ff91b4000000            call dword ptr [ecx+000000b4]
'0070a16e    dbe2                    fnclex
'0070a170    3bc6                    cmp eax, esi
'0070a172    7d11                    jge 70a185

If (var_83 < __vbaVarCat) Then
'0070a174    68b4000000              push 000000b4
'0070a179    6830314300              push 00433130
'0070a17e    8b4da4                  mov ecx, dword ptr [ebp-5c]
'0070a181    51                      push ecx
'0070a182    50                      push eax
'0070a183    ffd3                    call ebx
    
End If
'0070a185    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'0070a18b    898d80fdffff            mov dword ptr [ebp+fffffd80], ecx
'0070a191    c78584feffffac874300    mov dword ptr [ebp+fffffe84], 004387ac
'0070a19b    b808000000              mov eax, 00000008
'0070a1a0    89857cfeffff            mov dword ptr [ebp+fffffe7c], eax
'0070a1a6    8b11                    mov edx, dword ptr [ecx]
'0070a1a8    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'0070a1ae    51                      push ecx
'0070a1af    83ec10                  sub esp, 10
'0070a1b2    8bcc                    mov ecx, esp
'0070a1b4    8901                    mov dword ptr [ecx], eax
'0070a1b6    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'0070a1bc    894104                  mov dword ptr [ecx+04], eax
'0070a1bf    8b8584feffff            mov eax, dword ptr [ebp+fffffe84]
'0070a1c5    894108                  mov dword ptr [ecx+08], eax
'0070a1c8    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'0070a1ce    89410c                  mov dword ptr [ecx+0c], eax
'0070a1d1    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'0070a1d7    51                      push ecx
'0070a1d8    ff5230                  call dword ptr [edx+30]
'0070a1db    dbe2                    fnclex
'0070a1dd    3bc6                    cmp eax, esi
'0070a1df    7d11                    jge 70a1f2

If (0 < __vbaVarCat) Then
'0070a1e1    6a30                    push 30
'0070a1e3    68d8304300              push 004330d8
'0070a1e8    8b9580fdffff            mov edx, dword ptr [ebp+fffffd80]
'0070a1ee    52                      push edx
'0070a1ef    50                      push eax
'0070a1f0    ffd3                    call ebx
    
End If
'0070a1f2    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'0070a1f8    89b56cffffff            mov dword ptr [ebp+ffffff6c], esi
'0070a1fe    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'0070a204    c7853cffffff09000000    mov dword ptr [ebp+ffffff3c], 00000009
'0070a20e    89b554ffffff            mov dword ptr [ebp+ffffff54], esi
'0070a214    c7854cffffff02000000    mov dword ptr [ebp+ffffff4c], 00000002
'0070a21e    668b8598fdffff          mov ax, word ptr [ebp+fffffd98]
'0070a225    66898574feffff          mov word ptr [ebp+fffffe74], ax
'0070a22c    c7856cfeffff0b000000    mov dword ptr [ebp+fffffe6c], 0000000b
'0070a236    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'0070a23c    51                      push ecx
'0070a23d    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'0070a243    52                      push edx
'0070a244    8d856cfeffff            lea eax, dword ptr [ebp+fffffe6c]
'0070a24a    50                      push eax
'0070a24b    8d8d2cffffff            lea ecx, dword ptr [ebp+ffffff2c]
'0070a251    51                      push ecx

' *** Reference to "rtcImmediateIf"
'0070a252    ff154c124000            call dword ptr [0040124c]
'0070a258    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'0070a25e    52                      push edx

' *** Reference to "__vbaI2Var"
'0070a25f    ff150c124000            call dword ptr [0040120c]
'0070a265    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'0070a26b    50                      push eax
'0070a26c    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0070a272    51                      push ecx
'0070a273    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070a275    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_19)
'0070a27b    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'0070a281    52                      push edx
'0070a282    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'0070a288    50                      push eax
'0070a289    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'0070a28f    51                      push ecx
'0070a290    8d956cfeffff            lea edx, dword ptr [ebp+fffffe6c]
'0070a296    52                      push edx
'0070a297    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070a29d    50                      push eax
'0070a29e    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0070a2a0    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_88, var_346, var_89, var_90, var_93)
'0070a2a6    83c424                  add esp, 24
'0070a2a9    8b45a4                  mov eax, dword ptr [ebp-5c]
'0070a2ac    8b08                    mov ecx, dword ptr [eax]
'0070a2ae    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0070a2b4    52                      push edx
'0070a2b5    50                      push eax
'0070a2b6    ff91b4000000            call dword ptr [ecx+000000b4]
'0070a2bc    dbe2                    fnclex
'0070a2be    3bc6                    cmp eax, esi
'0070a2c0    7d11                    jge 70a2d3

If (var_83 < __vbaVarCat) Then
'0070a2c2    68b4000000              push 000000b4
'0070a2c7    6830314300              push 00433130
'0070a2cc    8b4da4                  mov ecx, dword ptr [ebp-5c]
'0070a2cf    51                      push ecx
'0070a2d0    50                      push eax
'0070a2d1    ffd3                    call ebx
    
End If
'0070a2d3    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'0070a2d9    898d8cfdffff            mov dword ptr [ebp+fffffd8c], ecx
'0070a2df    c78594feffff38784300    mov dword ptr [ebp+fffffe94], 00437838
'0070a2e9    b808000000              mov eax, 00000008
'0070a2ee    89858cfeffff            mov dword ptr [ebp+fffffe8c], eax
'0070a2f4    8b11                    mov edx, dword ptr [ecx]
'0070a2f6    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'0070a2fc    51                      push ecx
'0070a2fd    83ec10                  sub esp, 10
'0070a300    8bcc                    mov ecx, esp
'0070a302    8901                    mov dword ptr [ecx], eax
'0070a304    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'0070a30a    894104                  mov dword ptr [ecx+04], eax
'0070a30d    8b8594feffff            mov eax, dword ptr [ebp+fffffe94]
'0070a313    894108                  mov dword ptr [ecx+08], eax
'0070a316    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0070a31c    89410c                  mov dword ptr [ecx+0c], eax
'0070a31f    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'0070a325    51                      push ecx
'0070a326    ff5230                  call dword ptr [edx+30]
var_87.Description = 
'0070a329    dbe2                    fnclex
'0070a32b    3bc6                    cmp eax, esi
'0070a32d    7d11                    jge 70a340
'0070a32f    6a30                    push 30
'0070a331    68d8304300              push 004330d8
'0070a336    8b958cfdffff            mov edx, dword ptr [ebp+fffffd8c]
'0070a33c    52                      push edx
'0070a33d    50                      push eax
'0070a33e    ffd3                    call ebx
'0070a340    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'0070a346    89b574ffffff            mov dword ptr [ebp+ffffff74], esi
'0070a34c    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'0070a352    c7855cffffff09000000    mov dword ptr [ebp+ffffff5c], 00000009
'0070a35c    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070a362    50                      push eax

' *** Reference to "rtcIsNull"
'0070a363    ff1540114000            call dword ptr [00401140]
'0070a369    898598fdffff            mov dword ptr [ebp+fffffd98], eax
'0070a36f    8b45a4                  mov eax, dword ptr [ebp-5c]
'0070a372    8b08                    mov ecx, dword ptr [eax]
'0070a374    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0070a37a    52                      push edx
'0070a37b    50                      push eax
'0070a37c    ff91b4000000            call dword ptr [ecx+000000b4]
'0070a382    dbe2                    fnclex
'0070a384    3bc6                    cmp eax, esi
'0070a386    7d11                    jge 70a399

If (var_83 < __vbaVarCat) Then
'0070a388    68b4000000              push 000000b4
'0070a38d    6830314300              push 00433130
'0070a392    8b4da4                  mov ecx, dword ptr [ebp-5c]
'0070a395    51                      push ecx
'0070a396    50                      push eax
'0070a397    ffd3                    call ebx
    
End If
'0070a399    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'0070a39f    898d80fdffff            mov dword ptr [ebp+fffffd80], ecx
'0070a3a5    c78584feffff38784300    mov dword ptr [ebp+fffffe84], 00437838
'0070a3af    b808000000              mov eax, 00000008
'0070a3b4    89857cfeffff            mov dword ptr [ebp+fffffe7c], eax
'0070a3ba    8b11                    mov edx, dword ptr [ecx]
'0070a3bc    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'0070a3c2    51                      push ecx
'0070a3c3    83ec10                  sub esp, 10
'0070a3c6    8bcc                    mov ecx, esp
'0070a3c8    8901                    mov dword ptr [ecx], eax
'0070a3ca    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'0070a3d0    894104                  mov dword ptr [ecx+04], eax
'0070a3d3    8b8584feffff            mov eax, dword ptr [ebp+fffffe84]
'0070a3d9    894108                  mov dword ptr [ecx+08], eax
'0070a3dc    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'0070a3e2    89410c                  mov dword ptr [ecx+0c], eax
'0070a3e5    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'0070a3eb    51                      push ecx
'0070a3ec    ff5230                  call dword ptr [edx+30]
'0070a3ef    dbe2                    fnclex
'0070a3f1    3bc6                    cmp eax, esi
'0070a3f3    7d11                    jge 70a406

If (0 < __vbaVarCat) Then
'0070a3f5    6a30                    push 30
'0070a3f7    68d8304300              push 004330d8
'0070a3fc    8b9580fdffff            mov edx, dword ptr [ebp+fffffd80]
'0070a402    52                      push edx
'0070a403    50                      push eax
'0070a404    ffd3                    call ebx
    
End If
'0070a406    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'0070a40c    89b56cffffff            mov dword ptr [ebp+ffffff6c], esi
'0070a412    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'0070a418    c7853cffffff09000000    mov dword ptr [ebp+ffffff3c], 00000009
'0070a422    89b554ffffff            mov dword ptr [ebp+ffffff54], esi
'0070a428    c7854cffffff02000000    mov dword ptr [ebp+ffffff4c], 00000002
'0070a432    668b8598fdffff          mov ax, word ptr [ebp+fffffd98]
'0070a439    66898574feffff          mov word ptr [ebp+fffffe74], ax
'0070a440    c7856cfeffff0b000000    mov dword ptr [ebp+fffffe6c], 0000000b
'0070a44a    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'0070a450    51                      push ecx
'0070a451    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'0070a457    52                      push edx
'0070a458    8d856cfeffff            lea eax, dword ptr [ebp+fffffe6c]
'0070a45e    50                      push eax
'0070a45f    8d8d2cffffff            lea ecx, dword ptr [ebp+ffffff2c]
'0070a465    51                      push ecx

' *** Reference to "rtcImmediateIf"
'0070a466    ff154c124000            call dword ptr [0040124c]
'0070a46c    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'0070a472    52                      push edx

' *** Reference to "__vbaI2Var"
'0070a473    ff150c124000            call dword ptr [0040120c]
'0070a479    8945c8                  mov dword ptr [ebp-38], eax
'0070a47c    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'0070a482    50                      push eax
'0070a483    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0070a489    51                      push ecx
'0070a48a    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070a48c    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_19)
'0070a492    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'0070a498    52                      push edx
'0070a499    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'0070a49f    50                      push eax
'0070a4a0    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'0070a4a6    51                      push ecx
'0070a4a7    8d956cfeffff            lea edx, dword ptr [ebp+fffffe6c]
'0070a4ad    52                      push edx
'0070a4ae    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070a4b4    50                      push eax
'0070a4b5    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0070a4b7    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_88, var_346, var_89, var_90, var_93)
'0070a4bd    83c424                  add esp, 24
'0070a4c0    8b45a4                  mov eax, dword ptr [ebp-5c]
'0070a4c3    8b08                    mov ecx, dword ptr [eax]
'0070a4c5    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0070a4cb    52                      push edx
'0070a4cc    50                      push eax
'0070a4cd    ff91b4000000            call dword ptr [ecx+000000b4]
'0070a4d3    dbe2                    fnclex
'0070a4d5    3bc6                    cmp eax, esi
'0070a4d7    7d11                    jge 70a4ea

If (var_83 < __vbaVarCat) Then
'0070a4d9    68b4000000              push 000000b4
'0070a4de    6830314300              push 00433130
'0070a4e3    8b4da4                  mov ecx, dword ptr [ebp-5c]
'0070a4e6    51                      push ecx
'0070a4e7    50                      push eax
'0070a4e8    ffd3                    call ebx
    
End If
'0070a4ea    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'0070a4f0    898d8cfdffff            mov dword ptr [ebp+fffffd8c], ecx
'0070a4f6    c78594feffffc4d34300    mov dword ptr [ebp+fffffe94], 0043d3c4
'0070a500    b808000000              mov eax, 00000008
'0070a505    89858cfeffff            mov dword ptr [ebp+fffffe8c], eax
'0070a50b    8b11                    mov edx, dword ptr [ecx]
'0070a50d    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'0070a513    51                      push ecx
'0070a514    83ec10                  sub esp, 10
'0070a517    8bcc                    mov ecx, esp
'0070a519    8901                    mov dword ptr [ecx], eax
'0070a51b    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'0070a521    894104                  mov dword ptr [ecx+04], eax
'0070a524    8b8594feffff            mov eax, dword ptr [ebp+fffffe94]
'0070a52a    894108                  mov dword ptr [ecx+08], eax
'0070a52d    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0070a533    89410c                  mov dword ptr [ecx+0c], eax
'0070a536    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'0070a53c    51                      push ecx
'0070a53d    ff5230                  call dword ptr [edx+30]
var_87.Description = 
'0070a540    dbe2                    fnclex
'0070a542    3bc6                    cmp eax, esi
'0070a544    7d11                    jge 70a557
'0070a546    6a30                    push 30
'0070a548    68d8304300              push 004330d8
'0070a54d    8b958cfdffff            mov edx, dword ptr [ebp+fffffd8c]
'0070a553    52                      push edx
'0070a554    50                      push eax
'0070a555    ffd3                    call ebx
'0070a557    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'0070a55d    89b574ffffff            mov dword ptr [ebp+ffffff74], esi
'0070a563    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'0070a569    c7855cffffff09000000    mov dword ptr [ebp+ffffff5c], 00000009
'0070a573    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070a579    50                      push eax

' *** Reference to "rtcIsNull"
'0070a57a    ff1540114000            call dword ptr [00401140]
'0070a580    898598fdffff            mov dword ptr [ebp+fffffd98], eax
'0070a586    8b45a4                  mov eax, dword ptr [ebp-5c]
'0070a589    8b08                    mov ecx, dword ptr [eax]
'0070a58b    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0070a591    52                      push edx
'0070a592    50                      push eax
'0070a593    ff91b4000000            call dword ptr [ecx+000000b4]
'0070a599    dbe2                    fnclex
'0070a59b    3bc6                    cmp eax, esi
'0070a59d    7d11                    jge 70a5b0

If (var_83 < __vbaVarCat) Then
'0070a59f    68b4000000              push 000000b4
'0070a5a4    6830314300              push 00433130
'0070a5a9    8b4da4                  mov ecx, dword ptr [ebp-5c]
'0070a5ac    51                      push ecx
'0070a5ad    50                      push eax
'0070a5ae    ffd3                    call ebx
    
End If
'0070a5b0    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'0070a5b6    898d80fdffff            mov dword ptr [ebp+fffffd80], ecx
'0070a5bc    c78584feffffc4d34300    mov dword ptr [ebp+fffffe84], 0043d3c4
'0070a5c6    b808000000              mov eax, 00000008
'0070a5cb    89857cfeffff            mov dword ptr [ebp+fffffe7c], eax
'0070a5d1    8b11                    mov edx, dword ptr [ecx]
'0070a5d3    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'0070a5d9    51                      push ecx
'0070a5da    83ec10                  sub esp, 10
'0070a5dd    8bcc                    mov ecx, esp
'0070a5df    8901                    mov dword ptr [ecx], eax
'0070a5e1    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'0070a5e7    894104                  mov dword ptr [ecx+04], eax
'0070a5ea    8b8584feffff            mov eax, dword ptr [ebp+fffffe84]
'0070a5f0    894108                  mov dword ptr [ecx+08], eax
'0070a5f3    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'0070a5f9    89410c                  mov dword ptr [ecx+0c], eax
'0070a5fc    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'0070a602    51                      push ecx
'0070a603    ff5230                  call dword ptr [edx+30]
'0070a606    dbe2                    fnclex
'0070a608    3bc6                    cmp eax, esi
'0070a60a    7d11                    jge 70a61d

If (0 < __vbaVarCat) Then
'0070a60c    6a30                    push 30
'0070a60e    68d8304300              push 004330d8
'0070a613    8b9580fdffff            mov edx, dword ptr [ebp+fffffd80]
'0070a619    52                      push edx
'0070a61a    50                      push eax
'0070a61b    ffd3                    call ebx
    
End If
'0070a61d    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'0070a623    89b56cffffff            mov dword ptr [ebp+ffffff6c], esi
'0070a629    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'0070a62f    c7853cffffff09000000    mov dword ptr [ebp+ffffff3c], 00000009
'0070a639    89b554ffffff            mov dword ptr [ebp+ffffff54], esi
'0070a63f    c7854cffffff02000000    mov dword ptr [ebp+ffffff4c], 00000002
'0070a649    668b8598fdffff          mov ax, word ptr [ebp+fffffd98]
'0070a650    66898574feffff          mov word ptr [ebp+fffffe74], ax
'0070a657    c7856cfeffff0b000000    mov dword ptr [ebp+fffffe6c], 0000000b
'0070a661    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'0070a667    51                      push ecx
'0070a668    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'0070a66e    52                      push edx
'0070a66f    8d856cfeffff            lea eax, dword ptr [ebp+fffffe6c]
'0070a675    50                      push eax
'0070a676    8d8d2cffffff            lea ecx, dword ptr [ebp+ffffff2c]
'0070a67c    51                      push ecx

' *** Reference to "rtcImmediateIf"
'0070a67d    ff154c124000            call dword ptr [0040124c]
'0070a683    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'0070a689    52                      push edx

' *** Reference to "__vbaI2Var"
'0070a68a    ff150c124000            call dword ptr [0040120c]
'0070a690    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'0070a696    50                      push eax
'0070a697    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0070a69d    51                      push ecx
'0070a69e    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070a6a0    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_19)
'0070a6a6    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'0070a6ac    52                      push edx
'0070a6ad    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'0070a6b3    50                      push eax
'0070a6b4    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'0070a6ba    51                      push ecx
'0070a6bb    8d956cfeffff            lea edx, dword ptr [ebp+fffffe6c]
'0070a6c1    52                      push edx
'0070a6c2    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070a6c8    50                      push eax
'0070a6c9    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0070a6cb    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_88, var_346, var_89, var_90, var_93)
'0070a6d1    83c424                  add esp, 24
'0070a6d4    8b45c4                  mov eax, dword ptr [ebp-3c]
'0070a6d7    8b08                    mov ecx, dword ptr [eax]
'0070a6d9    8d9598fdffff            lea edx, dword ptr [ebp+fffffd98]
'0070a6df    52                      push edx
'0070a6e0    50                      push eax
'0070a6e1    ff515c                  call dword ptr [ecx+5c]
'0070a6e4    dbe2                    fnclex
'0070a6e6    3bc6                    cmp eax, esi
'0070a6e8    7d0e                    jge 70a6f8

If (var_9 < __vbaVarCat) Then
'0070a6ea    6a5c                    push 5c
'0070a6ec    6830314300              push 00433130
'0070a6f1    8b4dc4                  mov ecx, dword ptr [ebp-3c]
'0070a6f4    51                      push ecx
'0070a6f5    50                      push eax
'0070a6f6    ffd3                    call ebx
    
End If
'0070a6f8    6639b598fdffff          cmp word ptr [ebp+fffffd98], si
'0070a6ff    7407                    je 70a708

If (IsNull(__vbaVarCat) <> __vbaVarCat) Then
'0070a701    33f6                    xor esi, esi
    var_num8 = Empty
'0070a703    e918020000              jmp 70a920
    
End If
'0070a708    8b45c4                  mov eax, dword ptr [ebp-3c]
'0070a70b    8b10                    mov edx, dword ptr [eax]
'0070a70d    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0070a713    51                      push ecx
'0070a714    50                      push eax
'0070a715    ff92b4000000            call dword ptr [edx+000000b4]
'0070a71b    dbe2                    fnclex
'0070a71d    3bc6                    cmp eax, esi
'0070a71f    7d11                    jge 70a732

If (var_9 < __vbaVarCat) Then
'0070a721    68b4000000              push 000000b4
'0070a726    6830314300              push 00433130
'0070a72b    8b55c4                  mov edx, dword ptr [ebp-3c]
'0070a72e    52                      push edx
'0070a72f    50                      push eax
'0070a730    ffd3                    call ebx
    
End If
'0070a732    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'0070a738    898d8cfdffff            mov dword ptr [ebp+fffffd8c], ecx
'0070a73e    c78594fefffff8c04300    mov dword ptr [ebp+fffffe94], 0043c0f8
'0070a748    b808000000              mov eax, 00000008
'0070a74d    89858cfeffff            mov dword ptr [ebp+fffffe8c], eax
'0070a753    8b09                    mov ecx, dword ptr [ecx]
'0070a755    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'0070a75b    52                      push edx
'0070a75c    83ec10                  sub esp, 10
'0070a75f    8bd4                    mov edx, esp
'0070a761    8902                    mov dword ptr [edx], eax
'0070a763    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'0070a769    894204                  mov dword ptr [edx+04], eax
'0070a76c    8b8594feffff            mov eax, dword ptr [ebp+fffffe94]
'0070a772    894208                  mov dword ptr [edx+08], eax
'0070a775    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0070a77b    89420c                  mov dword ptr [edx+0c], eax
'0070a77e    8b9578ffffff            mov edx, dword ptr [ebp+ffffff78]
'0070a784    52                      push edx
'0070a785    ff5130                  call dword ptr [ecx+30]
var_87.Description = 
'0070a788    dbe2                    fnclex
'0070a78a    3bc6                    cmp eax, esi
'0070a78c    7d11                    jge 70a79f
'0070a78e    6a30                    push 30
'0070a790    68d8304300              push 004330d8
'0070a795    8b8d8cfdffff            mov ecx, dword ptr [ebp+fffffd8c]
'0070a79b    51                      push ecx
'0070a79c    50                      push eax
'0070a79d    ffd3                    call ebx
'0070a79f    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'0070a7a5    89b574ffffff            mov dword ptr [ebp+ffffff74], esi
'0070a7ab    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'0070a7b1    c7855cffffff09000000    mov dword ptr [ebp+ffffff5c], 00000009
'0070a7bb    8d955cffffff            lea edx, dword ptr [ebp+ffffff5c]
'0070a7c1    52                      push edx

' *** Reference to "rtcIsNull"
'0070a7c2    ff1540114000            call dword ptr [00401140]
'0070a7c8    898598fdffff            mov dword ptr [ebp+fffffd98], eax
'0070a7ce    8b45c4                  mov eax, dword ptr [ebp-3c]
'0070a7d1    8b08                    mov ecx, dword ptr [eax]
'0070a7d3    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0070a7d9    52                      push edx
'0070a7da    50                      push eax
'0070a7db    ff91b4000000            call dword ptr [ecx+000000b4]
'0070a7e1    dbe2                    fnclex
'0070a7e3    3bc6                    cmp eax, esi
'0070a7e5    7d11                    jge 70a7f8

If (var_9 < __vbaVarCat) Then
'0070a7e7    68b4000000              push 000000b4
'0070a7ec    6830314300              push 00433130
'0070a7f1    8b4dc4                  mov ecx, dword ptr [ebp-3c]
'0070a7f4    51                      push ecx
'0070a7f5    50                      push eax
'0070a7f6    ffd3                    call ebx
    
End If
'0070a7f8    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'0070a7fe    898d80fdffff            mov dword ptr [ebp+fffffd80], ecx
'0070a804    c78584fefffff8c04300    mov dword ptr [ebp+fffffe84], 0043c0f8
'0070a80e    b808000000              mov eax, 00000008
'0070a813    89857cfeffff            mov dword ptr [ebp+fffffe7c], eax
'0070a819    8b11                    mov edx, dword ptr [ecx]
'0070a81b    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'0070a821    51                      push ecx
'0070a822    83ec10                  sub esp, 10
'0070a825    8bcc                    mov ecx, esp
'0070a827    8901                    mov dword ptr [ecx], eax
'0070a829    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'0070a82f    894104                  mov dword ptr [ecx+04], eax
'0070a832    8b8584feffff            mov eax, dword ptr [ebp+fffffe84]
'0070a838    894108                  mov dword ptr [ecx+08], eax
'0070a83b    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'0070a841    89410c                  mov dword ptr [ecx+0c], eax
'0070a844    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'0070a84a    51                      push ecx
'0070a84b    ff5230                  call dword ptr [edx+30]
'0070a84e    dbe2                    fnclex
'0070a850    3bc6                    cmp eax, esi
'0070a852    7d11                    jge 70a865

If (0 < __vbaVarCat) Then
'0070a854    6a30                    push 30
'0070a856    68d8304300              push 004330d8
'0070a85b    8b9580fdffff            mov edx, dword ptr [ebp+fffffd80]
'0070a861    52                      push edx
'0070a862    50                      push eax
'0070a863    ffd3                    call ebx
    
End If
'0070a865    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'0070a86b    33c9                    xor ecx, ecx
var_num3 = Empty
'0070a86d    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'0070a873    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'0070a879    c7853cffffff09000000    mov dword ptr [ebp+ffffff3c], 00000009
'0070a883    898d54ffffff            mov dword ptr [ebp+ffffff54], ecx
'0070a889    c7854cffffff02000000    mov dword ptr [ebp+ffffff4c], 00000002
'0070a893    668b8598fdffff          mov ax, word ptr [ebp+fffffd98]
'0070a89a    66898574feffff          mov word ptr [ebp+fffffe74], ax
'0070a8a1    c7856cfeffff0b000000    mov dword ptr [ebp+fffffe6c], 0000000b
'0070a8ab    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'0070a8b1    51                      push ecx
'0070a8b2    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'0070a8b8    52                      push edx
'0070a8b9    8d856cfeffff            lea eax, dword ptr [ebp+fffffe6c]
'0070a8bf    50                      push eax
'0070a8c0    8d8d2cffffff            lea ecx, dword ptr [ebp+ffffff2c]
'0070a8c6    51                      push ecx

' *** Reference to "rtcImmediateIf"
'0070a8c7    ff154c124000            call dword ptr [0040124c]
'0070a8cd    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'0070a8d3    52                      push edx

' *** Reference to "__vbaI2Var"
'0070a8d4    ff150c124000            call dword ptr [0040120c]
'0070a8da    8bf0                    mov esi, eax
'0070a8dc    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'0070a8e2    50                      push eax
'0070a8e3    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0070a8e9    51                      push ecx
'0070a8ea    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070a8ec    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_19)
'0070a8f2    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'0070a8f8    52                      push edx
'0070a8f9    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'0070a8ff    50                      push eax
'0070a900    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'0070a906    51                      push ecx
'0070a907    8d956cfeffff            lea edx, dword ptr [ebp+fffffe6c]
'0070a90d    52                      push edx
'0070a90e    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070a914    50                      push eax
'0070a915    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0070a917    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_88, var_346, var_89, var_90, var_93)
'0070a91d    83c424                  add esp, 24
'0070a920    8b45b8                  mov eax, dword ptr [ebp-48]
'0070a923    8b08                    mov ecx, dword ptr [eax]
'0070a925    8d9598fdffff            lea edx, dword ptr [ebp+fffffd98]
'0070a92b    52                      push edx
'0070a92c    50                      push eax
'0070a92d    ff515c                  call dword ptr [ecx+5c]
'0070a930    dbe2                    fnclex
'0070a932    85c0                    test ax, ax
'0070a934    7d0e                    jge 70a944
'0070a936    6a5c                    push 5c
'0070a938    6830314300              push 00433130
'0070a93d    8b4db8                  mov ecx, dword ptr [ebp-48]
'0070a940    51                      push ecx
'0070a941    50                      push eax
'0070a942    ffd3                    call ebx
'0070a944    6683bd98fdffff00        cmp word ptr [ebp+fffffd98], 00
'0070a94c    7407                    je 70a955

If (IsNull(__vbaVarCat) <> 0) Then
'0070a94e    33c9                    xor ecx, ecx
    var_num3 = Empty
'0070a950    e920020000              jmp 70ab75
    
End If
'0070a955    8b45b8                  mov eax, dword ptr [ebp-48]
'0070a958    8b10                    mov edx, dword ptr [eax]
'0070a95a    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0070a960    51                      push ecx
'0070a961    50                      push eax
'0070a962    ff92b4000000            call dword ptr [edx+000000b4]
'0070a968    dbe2                    fnclex
'0070a96a    85c0                    test ax, ax
'0070a96c    7d11                    jge 70a97f
'0070a96e    68b4000000              push 000000b4
'0070a973    6830314300              push 00433130
'0070a978    8b55b8                  mov edx, dword ptr [ebp-48]
'0070a97b    52                      push edx
'0070a97c    50                      push eax
'0070a97d    ffd3                    call ebx
'0070a97f    8b8d78ffffff            mov ecx, dword ptr [ebp+ffffff78]
'0070a985    898d8cfdffff            mov dword ptr [ebp+fffffd8c], ecx
'0070a98b    c78594fefffff8c04300    mov dword ptr [ebp+fffffe94], 0043c0f8
'0070a995    b808000000              mov eax, 00000008
'0070a99a    89858cfeffff            mov dword ptr [ebp+fffffe8c], eax
'0070a9a0    8b09                    mov ecx, dword ptr [ecx]
'0070a9a2    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'0070a9a8    52                      push edx
'0070a9a9    83ec10                  sub esp, 10
'0070a9ac    8bd4                    mov edx, esp
'0070a9ae    8902                    mov dword ptr [edx], eax
'0070a9b0    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'0070a9b6    894204                  mov dword ptr [edx+04], eax
'0070a9b9    8b8594feffff            mov eax, dword ptr [ebp+fffffe94]
'0070a9bf    894208                  mov dword ptr [edx+08], eax
'0070a9c2    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0070a9c8    89420c                  mov dword ptr [edx+0c], eax
'0070a9cb    8b9578ffffff            mov edx, dword ptr [ebp+ffffff78]
'0070a9d1    52                      push edx
'0070a9d2    ff5130                  call dword ptr [ecx+30]
var_87.Description = 
'0070a9d5    dbe2                    fnclex
'0070a9d7    85c0                    test ax, ax
'0070a9d9    7d11                    jge 70a9ec
'0070a9db    6a30                    push 30
'0070a9dd    68d8304300              push 004330d8
'0070a9e2    8b8d8cfdffff            mov ecx, dword ptr [ebp+fffffd8c]
'0070a9e8    51                      push ecx
'0070a9e9    50                      push eax
'0070a9ea    ffd3                    call ebx
'0070a9ec    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'0070a9f2    c78574ffffff00000000    mov dword ptr [ebp+ffffff74], 00000000
'0070a9fc    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'0070aa02    c7855cffffff09000000    mov dword ptr [ebp+ffffff5c], 00000009
'0070aa0c    8d955cffffff            lea edx, dword ptr [ebp+ffffff5c]
'0070aa12    52                      push edx

' *** Reference to "rtcIsNull"
'0070aa13    ff1540114000            call dword ptr [00401140]
'0070aa19    898598fdffff            mov dword ptr [ebp+fffffd98], eax
'0070aa1f    8b45b8                  mov eax, dword ptr [ebp-48]
'0070aa22    8b08                    mov ecx, dword ptr [eax]
'0070aa24    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0070aa2a    52                      push edx
'0070aa2b    50                      push eax
'0070aa2c    ff91b4000000            call dword ptr [ecx+000000b4]
'0070aa32    dbe2                    fnclex
'0070aa34    85c0                    test ax, ax
'0070aa36    7d11                    jge 70aa49
'0070aa38    68b4000000              push 000000b4
'0070aa3d    6830314300              push 00433130
'0070aa42    8b4db8                  mov ecx, dword ptr [ebp-48]
'0070aa45    51                      push ecx
'0070aa46    50                      push eax
'0070aa47    ffd3                    call ebx
'0070aa49    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'0070aa4f    898d80fdffff            mov dword ptr [ebp+fffffd80], ecx
'0070aa55    c78584fefffff8c04300    mov dword ptr [ebp+fffffe84], 0043c0f8
'0070aa5f    b808000000              mov eax, 00000008
'0070aa64    89857cfeffff            mov dword ptr [ebp+fffffe7c], eax
'0070aa6a    8b11                    mov edx, dword ptr [ecx]
'0070aa6c    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'0070aa72    51                      push ecx
'0070aa73    83ec10                  sub esp, 10
'0070aa76    8bcc                    mov ecx, esp
'0070aa78    8901                    mov dword ptr [ecx], eax
'0070aa7a    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'0070aa80    894104                  mov dword ptr [ecx+04], eax
'0070aa83    8b8584feffff            mov eax, dword ptr [ebp+fffffe84]
'0070aa89    894108                  mov dword ptr [ecx+08], eax
'0070aa8c    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'0070aa92    89410c                  mov dword ptr [ecx+0c], eax
'0070aa95    8b8d70ffffff            mov ecx, dword ptr [ebp+ffffff70]
'0070aa9b    51                      push ecx
'0070aa9c    ff5230                  call dword ptr [edx+30]
'0070aa9f    dbe2                    fnclex
'0070aaa1    85c0                    test ax, ax
'0070aaa3    7d11                    jge 70aab6
'0070aaa5    6a30                    push 30
'0070aaa7    68d8304300              push 004330d8
'0070aaac    8b9580fdffff            mov edx, dword ptr [ebp+fffffd80]
'0070aab2    52                      push edx
'0070aab3    50                      push eax
'0070aab4    ffd3                    call ebx
'0070aab6    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'0070aabc    33c9                    xor ecx, ecx
var_num3 = Empty
'0070aabe    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'0070aac4    898544ffffff            mov dword ptr [ebp+ffffff44], eax
'0070aaca    c7853cffffff09000000    mov dword ptr [ebp+ffffff3c], 00000009
'0070aad4    898d54ffffff            mov dword ptr [ebp+ffffff54], ecx
'0070aada    c7854cffffff02000000    mov dword ptr [ebp+ffffff4c], 00000002
'0070aae4    668b8598fdffff          mov ax, word ptr [ebp+fffffd98]
'0070aaeb    66898574feffff          mov word ptr [ebp+fffffe74], ax
'0070aaf2    c7856cfeffff0b000000    mov dword ptr [ebp+fffffe6c], 0000000b
'0070aafc    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'0070ab02    51                      push ecx
'0070ab03    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'0070ab09    52                      push edx
'0070ab0a    8d856cfeffff            lea eax, dword ptr [ebp+fffffe6c]
'0070ab10    50                      push eax
'0070ab11    8d8d2cffffff            lea ecx, dword ptr [ebp+ffffff2c]
'0070ab17    51                      push ecx

' *** Reference to "rtcImmediateIf"
'0070ab18    ff154c124000            call dword ptr [0040124c]
'0070ab1e    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'0070ab24    52                      push edx

' *** Reference to "__vbaI2Var"
'0070ab25    ff150c124000            call dword ptr [0040120c]
'0070ab2b    8945ac                  mov dword ptr [ebp-54], eax
'0070ab2e    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'0070ab34    50                      push eax
'0070ab35    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0070ab3b    51                      push ecx
'0070ab3c    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070ab3e    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_19)
'0070ab44    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'0070ab4a    52                      push edx
'0070ab4b    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'0070ab51    50                      push eax
'0070ab52    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'0070ab58    51                      push ecx
'0070ab59    8d956cfeffff            lea edx, dword ptr [ebp+fffffe6c]
'0070ab5f    52                      push edx
'0070ab60    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070ab66    50                      push eax
'0070ab67    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0070ab69    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_88, var_346, var_89, var_90, var_93)
'0070ab6f    83c424                  add esp, 24
'0070ab72    8b4dac                  mov ecx, dword ptr [ebp-54]
'0070ab75    66037d9c                add di, word ptr [ebp-64]
var_num7 = CInt(IIf(IsNull(__vbaVarCat), __vbaVarCat, __vbaVarCat)) + CInt(IIf(IsNull(var_91), __vbaVarCat, 0))
'0070ab79    0f80c80f0000            jo 70bb47
'0070ab7f    66037ddc                add di, word ptr [ebp-24]
var_num7 = var_num7 + CInt(IIf(IsNull(__vbaVarCat), __vbaVarCat, __vbaVarCat))
'0070ab83    0f80be0f0000            jo 70bb47
'0070ab89    66037dd4                add di, word ptr [ebp-2c]
var_num7 = var_num7 + CInt(IIf(IsNull(__vbaVarCat), __vbaVarCat, __vbaVarCat))
'0070ab8d    0f80b40f0000            jo 70bb47
'0070ab93    66037dd0                add di, word ptr [ebp-30]
var_num7 = var_num7 + CInt(IIf(IsNull(__vbaVarCat), __vbaVarCat, __vbaVarCat))
'0070ab97    0f80aa0f0000            jo 70bb47
'0070ab9d    66037dcc                add di, word ptr [ebp-34]
var_num7 = var_num7 + CInt(IIf(IsNull(__vbaVarCat), __vbaVarCat, __vbaVarCat))
'0070aba1    0f80a00f0000            jo 70bb47
'0070aba7    66037dc8                add di, word ptr [ebp-38]
var_num7 = var_num7 + CInt(IIf(IsNull(__vbaVarCat), __vbaVarCat, __vbaVarCat))
'0070abab    0f80960f0000            jo 70bb47
'0070abb1    33c0                    xor eax, eax
'0070abb3    6603f8                  add di, ax
var_num7 = var_num7 + 0
'0070abb6    0f808b0f0000            jo 70bb47
'0070abbc    897dbc                  mov dword ptr [ebp-44], edi
'0070abbf    6603f7                  add si, di
var_num8 = CInt(IIf(IsNull(__vbaVarCat), __vbaVarCat, __vbaVarCat)) + var_num7
'0070abc2    0f807f0f0000            jo 70bb47
'0070abc8    6603f1                  add si, cx
var_num8 = var_num8 + CInt(IIf(IsNull(__vbaVarCat), __vbaVarCat, __vbaVarCat))
'0070abcb    0f80760f0000            jo 70bb47
'0070abd1    8b45d8                  mov eax, dword ptr [ebp-28]
'0070abd4    8b08                    mov ecx, dword ptr [eax]
'0070abd6    8d9598fdffff            lea edx, dword ptr [ebp+fffffd98]
'0070abdc    52                      push edx
'0070abdd    50                      push eax
'0070abde    ff5134                  call dword ptr [ecx+34]
'0070abe1    dbe2                    fnclex
'0070abe3    85c0                    test ax, ax
'0070abe5    7d0e                    jge 70abf5
'0070abe7    6a34                    push 34
'0070abe9    6830314300              push 00433130
'0070abee    8b4dd8                  mov ecx, dword ptr [ebp-28]
'0070abf1    51                      push ecx
'0070abf2    50                      push eax
'0070abf3    ffd3                    call ebx
'0070abf5    6683bd98fdffff00        cmp word ptr [ebp+fffffd98], 00
'0070abfd    0f854a010000            jne 70ad4d

Do While (IsNull(__vbaVarCat) = 0)
'0070ac03    8b45d8                  mov eax, dword ptr [ebp-28]
'0070ac06    8b10                    mov edx, dword ptr [eax]
'0070ac08    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0070ac0e    51                      push ecx
'0070ac0f    50                      push eax
'0070ac10    ff92b4000000            call dword ptr [edx+000000b4]
'0070ac16    dbe2                    fnclex
'0070ac18    85c0                    test ax, ax
'0070ac1a    7d11                    jge 70ac2d
'0070ac1c    68b4000000              push 000000b4
'0070ac21    6830314300              push 00433130
'0070ac26    8b55d8                  mov edx, dword ptr [ebp-28]
'0070ac29    52                      push edx
'0070ac2a    50                      push eax
'0070ac2b    ffd3                    call ebx
'0070ac2d    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'0070ac33    89858cfdffff            mov dword ptr [ebp+fffffd8c], eax
'0070ac39    b9049e4300              mov ecx, 00439e04
'0070ac3e    898d94feffff            mov dword ptr [ebp+fffffe94], ecx
'0070ac44    ba08000000              mov edx, 00000008
'0070ac49    89958cfeffff            mov dword ptr [ebp+fffffe8c], edx
'0070ac4f    8b38                    mov edi, dword ptr [eax]
'0070ac51    8d9d74ffffff            lea ebx, dword ptr [ebp+ffffff74]
'0070ac57    53                      push ebx
'0070ac58    83ec10                  sub esp, 10
'0070ac5b    8bdc                    mov ebx, esp
'0070ac5d    8913                    mov dword ptr [ebx], edx
'0070ac5f    8b9590feffff            mov edx, dword ptr [ebp+fffffe90]
'0070ac65    895304                  mov dword ptr [ebx+04], edx
'0070ac68    894b08                  mov dword ptr [ebx+08], ecx
'0070ac6b    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'0070ac71    894b0c                  mov dword ptr [ebx+0c], ecx
'0070ac74    50                      push eax
'0070ac75    ff5730                  call dword ptr [edi+30]
    var_87.Description = 
'0070ac78    dbe2                    fnclex
'0070ac7a    85c0                    test ax, ax
'0070ac7c    7d19                    jge 70ac97
'0070ac7e    6a30                    push 30
'0070ac80    68d8304300              push 004330d8
'0070ac85    8b958cfdffff            mov edx, dword ptr [ebp+fffffd8c]
'0070ac8b    52                      push edx
'0070ac8c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070ac8d    8b1d80104000            mov ebx, dword ptr [00401080]
'0070ac93    ffd3                    call ebx
'0070ac95    eb06                    jmp 70ac9d

' *** Reference to "__vbaHresultCheckObj"
'0070ac97    8b1d80104000            mov ebx, dword ptr [00401080]
'0070ac9d    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'0070aca3    8bf8                    mov edi, eax
'0070aca5    8b08                    mov ecx, dword ptr [eax]
'0070aca7    8d955cffffff            lea edx, dword ptr [ebp+ffffff5c]
'0070acad    52                      push edx
'0070acae    50                      push eax
'0070acaf    ff5144                  call dword ptr [ecx+44]
'0070acb2    dbe2                    fnclex
'0070acb4    85c0                    test ax, ax
'0070acb6    7d0b                    jge 70acc3
'0070acb8    6a44                    push 44
'0070acba    6880304300              push 00433080
'0070acbf    57                      push edi
'0070acc0    50                      push eax
'0070acc1    ffd3                    call ebx
'0070acc3    c78584feffff98d14200    mov dword ptr [ebp+fffffe84], 0042d198
'0070accd    c7857cfeffff08800000    mov dword ptr [ebp+fffffe7c], 00008008
'0070acd7    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070acdd    50                      push eax
'0070acde    8d8d7cfeffff            lea ecx, dword ptr [ebp+fffffe7c]
'0070ace4    51                      push ecx

' *** Reference to "__vbaVarTstEq"
'0070ace5    ff153c114000            call dword ptr [0040113c]
'0070aceb    8bf8                    mov edi, eax
'0070aced    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'0070acf3    52                      push edx
'0070acf4    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0070acfa    50                      push eax
'0070acfb    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070acfd    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_87, var_91)
'0070ad03    83c40c                  add esp, 0c
'0070ad06    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]

' *** Reference to "__vbaFreeVar"
'0070ad0c    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_88)
'0070ad12    6685ff                  test edi, edi
'0070ad15    740a                    je 70ad21
    
    If (    ((__vbaVarCat) = ("Réduction de l'ajustement de niveau_1"))) Then
'0070ad17    6683ee01                sub si, 01
    var_num8 = var_num8 - 1
'0070ad1b    0f80260e0000            jo 70bb47
    
End If
'0070ad21    8b45d8                  mov eax, dword ptr [ebp-28]
'0070ad24    8b08                    mov ecx, dword ptr [eax]
'0070ad26    50                      push eax
'0070ad27    ff91ec000000            call dword ptr [ecx+000000ec]
'0070ad2d    dbe2                    fnclex
'0070ad2f    85c0                    test ax, ax
'0070ad31    0f8d9afeffff            jge 70abd1
'0070ad37    68ec000000              push 000000ec
'0070ad3c    6830314300              push 00433130
'0070ad41    8b55d8                  mov edx, dword ptr [ebp-28]
'0070ad44    52                      push edx
'0070ad45    50                      push eax
'0070ad46    ffd3                    call ebx
'0070ad48    e984feffff              jmp 70abd1

'ERROR: Two many next close:
Loop
'0070ad4d    8b45a4                  mov eax, dword ptr [ebp-5c]
'0070ad50    8b08                    mov ecx, dword ptr [eax]
'0070ad52    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0070ad58    52                      push edx
'0070ad59    50                      push eax
'0070ad5a    ff91b4000000            call dword ptr [ecx+000000b4]
'0070ad60    dbe2                    fnclex
'0070ad62    85c0                    test ax, ax
'0070ad64    7d11                    jge 70ad77
'0070ad66    68b4000000              push 000000b4
'0070ad6b    6830314300              push 00433130
'0070ad70    8b4da4                  mov ecx, dword ptr [ebp-5c]
'0070ad73    51                      push ecx
'0070ad74    50                      push eax
'0070ad75    ffd3                    call ebx
'0070ad77    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'0070ad7d    89858cfdffff            mov dword ptr [ebp+fffffd8c], eax
'0070ad83    c78594feffff84af4300    mov dword ptr [ebp+fffffe94], 0043af84
'0070ad8d    b908000000              mov ecx, 00000008
'0070ad92    898d8cfeffff            mov dword ptr [ebp+fffffe8c], ecx
'0070ad98    8b10                    mov edx, dword ptr [eax]
'0070ad9a    8dbd74ffffff            lea edi, dword ptr [ebp+ffffff74]
'0070ada0    57                      push edi
'0070ada1    83ec10                  sub esp, 10
'0070ada4    8bfc                    mov edi, esp
'0070ada6    890f                    mov dword ptr [edi], ecx
'0070ada8    8b8d90feffff            mov ecx, dword ptr [ebp+fffffe90]
'0070adae    894f04                  mov dword ptr [edi+04], ecx
'0070adb1    8b8d94feffff            mov ecx, dword ptr [ebp+fffffe94]
'0070adb7    894f08                  mov dword ptr [edi+08], ecx
'0070adba    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'0070adc0    894f0c                  mov dword ptr [edi+0c], ecx
'0070adc3    50                      push eax
'0070adc4    ff5230                  call dword ptr [edx+30]
var_87.Description = 
'0070adc7    dbe2                    fnclex
'0070adc9    85c0                    test ax, ax
'0070adcb    7d11                    jge 70adde
'0070adcd    6a30                    push 30
'0070adcf    68d8304300              push 004330d8
'0070add4    8b958cfdffff            mov edx, dword ptr [ebp+fffffd8c]
'0070adda    52                      push edx
'0070addb    50                      push eax
'0070addc    ffd3                    call ebx
'0070adde    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'0070ade4    8bf8                    mov edi, eax
'0070ade6    8b08                    mov ecx, dword ptr [eax]
'0070ade8    8d955cffffff            lea edx, dword ptr [ebp+ffffff5c]
'0070adee    52                      push edx
'0070adef    50                      push eax
'0070adf0    ff5144                  call dword ptr [ecx+44]
'0070adf3    dbe2                    fnclex
'0070adf5    85c0                    test ax, ax
'0070adf7    7d0b                    jge 70ae04
'0070adf9    6a44                    push 44
'0070adfb    6880304300              push 00433080
'0070ae00    57                      push edi
'0070ae01    50                      push eax
'0070ae02    ffd3                    call ebx
'0070ae04    56                      push esi

' *** Reference to "__vbaStrI2"
'0070ae05    ff1510104000            call dword ptr [00401010]
'0070ae0b    8bd0                    mov edx, eax
'0070ae0d    8d4d94                  lea ecx, dword ptr [ebp-6c]

' *** Reference to "__vbaStrMove"
'0070ae10    ff15d0124000            call dword ptr [004012d0]
'0070ae16    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'0070ae17    ff1510134000            call dword ptr [00401310]

' *** Reference to "__vbaFpI2"
'0070ae1d    ff15a0124000            call dword ptr [004012a0]
var_num1 = CLng(Val(CStr(var_num8)))
'0070ae23    898598fdffff            mov dword ptr [ebp+fffffd98], eax
'0070ae29    8d8598fdffff            lea eax, dword ptr [ebp+fffffd98]
'0070ae2f    50                      push eax
'0070ae30    e89b73deff              call 4f21d0
Call sub_4F21D0()
'0070ae35    898584feffff            mov dword ptr [ebp+fffffe84], eax
'0070ae3b    c7857cfeffff03800000    mov dword ptr [ebp+fffffe7c], 00008003
'0070ae45    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]
'0070ae4b    51                      push ecx
'0070ae4c    8d957cfeffff            lea edx, dword ptr [ebp+fffffe7c]
'0070ae52    52                      push edx

' *** Reference to "__vbaVarTstGt"
'0070ae53    ff1504104000            call dword ptr [00401004]
'0070ae59    8bf8                    mov edi, eax
'0070ae5b    8d4d94                  lea ecx, dword ptr [ebp-6c]

' *** Reference to "__vbaFreeStr"
'0070ae5e    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_121)
'0070ae64    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'0070ae6a    50                      push eax
'0070ae6b    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0070ae71    51                      push ecx
'0070ae72    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070ae74    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_91)
'0070ae7a    83c40c                  add esp, 0c
'0070ae7d    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]

' *** Reference to "__vbaFreeVar"
'0070ae83    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_88)
'0070ae89    6685ff                  test edi, edi
'0070ae8c    0f84ee000000            je 70af80

If (((__vbaVarCat) > (-872))) Then
'0070ae92    8b45a4                  mov eax, dword ptr [ebp-5c]
'0070ae95    8b10                    mov edx, dword ptr [eax]
'0070ae97    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0070ae9d    51                      push ecx
'0070ae9e    50                      push eax
'0070ae9f    ff92b4000000            call dword ptr [edx+000000b4]
'0070aea5    dbe2                    fnclex
'0070aea7    85c0                    test ax, ax
'0070aea9    7d11                    jge 70aebc
'0070aeab    68b4000000              push 000000b4
'0070aeb0    6830314300              push 00433130
'0070aeb5    8b55a4                  mov edx, dword ptr [ebp-5c]
'0070aeb8    52                      push edx
'0070aeb9    50                      push eax
'0070aeba    ffd3                    call ebx
'0070aebc    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'0070aec2    89858cfdffff            mov dword ptr [ebp+fffffd8c], eax
'0070aec8    c78594feffff84af4300    mov dword ptr [ebp+fffffe94], 0043af84
'0070aed2    b908000000              mov ecx, 00000008
'0070aed7    898d8cfeffff            mov dword ptr [ebp+fffffe8c], ecx
'0070aedd    8b10                    mov edx, dword ptr [eax]
'0070aedf    8dbd74ffffff            lea edi, dword ptr [ebp+ffffff74]
'0070aee5    57                      push edi
'0070aee6    83ec10                  sub esp, 10
'0070aee9    8bfc                    mov edi, esp
'0070aeeb    890f                    mov dword ptr [edi], ecx
'0070aeed    8b8d90feffff            mov ecx, dword ptr [ebp+fffffe90]
'0070aef3    894f04                  mov dword ptr [edi+04], ecx
'0070aef6    8b8d94feffff            mov ecx, dword ptr [ebp+fffffe94]
'0070aefc    894f08                  mov dword ptr [edi+08], ecx
'0070aeff    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'0070af05    894f0c                  mov dword ptr [edi+0c], ecx
'0070af08    50                      push eax
'0070af09    ff5230                  call dword ptr [edx+30]
    var_87.Description = 
'0070af0c    dbe2                    fnclex
'0070af0e    85c0                    test ax, ax
'0070af10    7d11                    jge 70af23
'0070af12    6a30                    push 30
'0070af14    68d8304300              push 004330d8
'0070af19    8b958cfdffff            mov edx, dword ptr [ebp+fffffd8c]
'0070af1f    52                      push edx
'0070af20    50                      push eax
'0070af21    ffd3                    call ebx
'0070af23    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'0070af29    8bf8                    mov edi, eax
'0070af2b    8b08                    mov ecx, dword ptr [eax]
'0070af2d    8d955cffffff            lea edx, dword ptr [ebp+ffffff5c]
'0070af33    52                      push edx
'0070af34    50                      push eax
'0070af35    ff5144                  call dword ptr [ecx+44]
'0070af38    dbe2                    fnclex
'0070af3a    85c0                    test ax, ax
'0070af3c    7d0b                    jge 70af49
'0070af3e    6a44                    push 44
'0070af40    6880304300              push 00433080
'0070af45    57                      push edi
'0070af46    50                      push eax
'0070af47    ffd3                    call ebx
'0070af49    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070af4f    50                      push eax

' *** Reference to "__vbaI4Var"
'0070af50    ff157c124000            call dword ptr [0040127c]
'0070af56    8945a0                  mov dword ptr [ebp-60], eax
'0070af59    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'0070af5f    51                      push ecx
'0070af60    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0070af66    52                      push edx
'0070af67    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070af69    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_87, var_91)
'0070af6f    83c40c                  add esp, 0c
'0070af72    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]

' *** Reference to "__vbaFreeVar"
'0070af78    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_88)
'0070af7e    eb3d                    jmp 70afbd
    
End If
'0070af80    56                      push esi

' *** Reference to "__vbaStrI2"
'0070af81    ff1510104000            call dword ptr [00401010]
'0070af87    8bd0                    mov edx, eax
'0070af89    8d4d94                  lea ecx, dword ptr [ebp-6c]

' *** Reference to "__vbaStrMove"
'0070af8c    ff15d0124000            call dword ptr [004012d0]
'0070af92    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'0070af93    ff1510134000            call dword ptr [00401310]

' *** Reference to "__vbaFpI2"
'0070af99    ff15a0124000            call dword ptr [004012a0]
var_num1 = CLng(Val(CStr(var_num8)))
'0070af9f    898598fdffff            mov dword ptr [ebp+fffffd98], eax
'0070afa5    8d8598fdffff            lea eax, dword ptr [ebp+fffffd98]
'0070afab    50                      push eax
'0070afac    e81f72deff              call 4f21d0
Call sub_4F21D0()
'0070afb1    8945a0                  mov dword ptr [ebp-60], eax
'0070afb4    8d4d94                  lea ecx, dword ptr [ebp-6c]

' *** Reference to "__vbaFreeStr"
'0070afb7    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_121)
'0070afbd    8b45a4                  mov eax, dword ptr [ebp-5c]
'0070afc0    8b08                    mov ecx, dword ptr [eax]
'0070afc2    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0070afc8    52                      push edx
'0070afc9    50                      push eax
'0070afca    ff91b4000000            call dword ptr [ecx+000000b4]
'0070afd0    dbe2                    fnclex
'0070afd2    85c0                    test ax, ax
'0070afd4    7d11                    jge 70afe7
'0070afd6    68b4000000              push 000000b4
'0070afdb    6830314300              push 00433130
'0070afe0    8b4da4                  mov ecx, dword ptr [ebp-5c]
'0070afe3    51                      push ecx
'0070afe4    50                      push eax
'0070afe5    ffd3                    call ebx
'0070afe7    8b8578ffffff            mov eax, dword ptr [ebp+ffffff78]
'0070afed    89858cfdffff            mov dword ptr [ebp+fffffd8c], eax
'0070aff3    c78594feffff64b14300    mov dword ptr [ebp+fffffe94], 0043b164
'0070affd    b908000000              mov ecx, 00000008
'0070b002    898d8cfeffff            mov dword ptr [ebp+fffffe8c], ecx
'0070b008    8b10                    mov edx, dword ptr [eax]
'0070b00a    8dbd74ffffff            lea edi, dword ptr [ebp+ffffff74]
'0070b010    57                      push edi
'0070b011    83ec10                  sub esp, 10
'0070b014    8bfc                    mov edi, esp
'0070b016    890f                    mov dword ptr [edi], ecx
'0070b018    8b8d90feffff            mov ecx, dword ptr [ebp+fffffe90]
'0070b01e    894f04                  mov dword ptr [edi+04], ecx
'0070b021    8b8d94feffff            mov ecx, dword ptr [ebp+fffffe94]
'0070b027    894f08                  mov dword ptr [edi+08], ecx
'0070b02a    8b8d98feffff            mov ecx, dword ptr [ebp+fffffe98]
'0070b030    894f0c                  mov dword ptr [edi+0c], ecx
'0070b033    50                      push eax
'0070b034    ff5230                  call dword ptr [edx+30]
var_87.Description = 
'0070b037    dbe2                    fnclex
'0070b039    85c0                    test ax, ax
'0070b03b    7d11                    jge 70b04e
'0070b03d    6a30                    push 30
'0070b03f    68d8304300              push 004330d8
'0070b044    8b958cfdffff            mov edx, dword ptr [ebp+fffffd8c]
'0070b04a    52                      push edx
'0070b04b    50                      push eax
'0070b04c    ffd3                    call ebx
'0070b04e    8b8574ffffff            mov eax, dword ptr [ebp+ffffff74]
'0070b054    8bf8                    mov edi, eax
'0070b056    8b08                    mov ecx, dword ptr [eax]
'0070b058    8d955cffffff            lea edx, dword ptr [ebp+ffffff5c]
'0070b05e    52                      push edx
'0070b05f    50                      push eax
'0070b060    ff5144                  call dword ptr [ecx+44]
'0070b063    dbe2                    fnclex
'0070b065    85c0                    test ax, ax
'0070b067    7d0b                    jge 70b074
'0070b069    6a44                    push 44
'0070b06b    6880304300              push 00433080
'0070b070    57                      push edi
'0070b071    50                      push eax
'0070b072    ffd3                    call ebx
'0070b074    bf44ed4300              mov edi, 0043ed44
'0070b079    89bd84feffff            mov dword ptr [ebp+fffffe84], edi
'0070b07f    bb08000000              mov ebx, 00000008
'0070b084    899d7cfeffff            mov dword ptr [ebp+fffffe7c], ebx
'0070b08a    668b45bc                mov ax, word ptr [ebp-44]
'0070b08e    66898574feffff          mov word ptr [ebp+fffffe74], ax
'0070b095    b802000000              mov eax, 00000002
'0070b09a    89856cfeffff            mov dword ptr [ebp+fffffe6c], eax
'0070b0a0    89bd64feffff            mov dword ptr [ebp+fffffe64], edi
'0070b0a6    899d5cfeffff            mov dword ptr [ebp+fffffe5c], ebx
'0070b0ac    6689b554feffff          mov word ptr [ebp+fffffe54], si
'0070b0b3    89854cfeffff            mov dword ptr [ebp+fffffe4c], eax
'0070b0b9    89bd44feffff            mov dword ptr [ebp+fffffe44], edi
'0070b0bf    899d3cfeffff            mov dword ptr [ebp+fffffe3c], ebx
'0070b0c5    8b4da0                  mov ecx, dword ptr [ebp-60]
'0070b0c8    898d34feffff            mov dword ptr [ebp+fffffe34], ecx
'0070b0ce    c7852cfeffff03000000    mov dword ptr [ebp+fffffe2c], 00000003
'0070b0d8    89bd24feffff            mov dword ptr [ebp+fffffe24], edi
'0070b0de    899d1cfeffff            mov dword ptr [ebp+fffffe1c], ebx
'0070b0e4    6683c601                add si, 01
var_num8 = var_num8 + 1
'0070b0e8    0f80590a0000            jo 70bb47
'0070b0ee    89b598fdffff            mov dword ptr [ebp+fffffd98], esi
'0070b0f4    8d9598fdffff            lea edx, dword ptr [ebp+fffffd98]
'0070b0fa    52                      push edx
'0070b0fb    e8d070deff              call 4f21d0
Call sub_4F21D0()
'0070b100    898514feffff            mov dword ptr [ebp+fffffe14], eax
'0070b106    c7850cfeffff03000000    mov dword ptr [ebp+fffffe0c], 00000003
'0070b110    89bd04feffff            mov dword ptr [ebp+fffffe04], edi
'0070b116    899dfcfdffff            mov dword ptr [ebp+fffffdfc], ebx
'0070b11c    89bdf4fdffff            mov dword ptr [ebp+fffffdf4], edi
'0070b122    899decfdffff            mov dword ptr [ebp+fffffdec], ebx
'0070b128    c785e4fdffff01000000    mov dword ptr [ebp+fffffde4], 00000001
'0070b132    c785dcfdffff02000000    mov dword ptr [ebp+fffffddc], 00000002
'0070b13c    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070b142    50                      push eax
'0070b143    8d8d7cfeffff            lea ecx, dword ptr [ebp+fffffe7c]
'0070b149    51                      push ecx
'0070b14a    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'0070b150    52                      push edx

' *** Reference to "__vbaVarCat"
'0070b151    8b3508124000            mov esi, dword ptr [00401208]
'0070b157    ffd6                    call esi
'0070b159    50                      push eax
'0070b15a    8d856cfeffff            lea eax, dword ptr [ebp+fffffe6c]
'0070b160    50                      push eax
'0070b161    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'0070b167    51                      push ecx
'0070b168    ffd6                    call esi
'0070b16a    50                      push eax
'0070b16b    8d955cfeffff            lea edx, dword ptr [ebp+fffffe5c]
'0070b171    52                      push edx
'0070b172    8d852cffffff            lea eax, dword ptr [ebp+ffffff2c]
'0070b178    50                      push eax
'0070b179    ffd6                    call esi
'0070b17b    50                      push eax
'0070b17c    8d8d4cfeffff            lea ecx, dword ptr [ebp+fffffe4c]
'0070b182    51                      push ecx
'0070b183    8d951cffffff            lea edx, dword ptr [ebp+ffffff1c]
'0070b189    52                      push edx
'0070b18a    ffd6                    call esi
'0070b18c    50                      push eax
'0070b18d    8d853cfeffff            lea eax, dword ptr [ebp+fffffe3c]
'0070b193    50                      push eax
'0070b194    8d8d0cffffff            lea ecx, dword ptr [ebp+ffffff0c]
'0070b19a    51                      push ecx
'0070b19b    ffd6                    call esi
'0070b19d    50                      push eax
'0070b19e    8d952cfeffff            lea edx, dword ptr [ebp+fffffe2c]
'0070b1a4    52                      push edx
'0070b1a5    8d85fcfeffff            lea eax, dword ptr [ebp+fffffefc]
'0070b1ab    50                      push eax
'0070b1ac    ffd6                    call esi
'0070b1ae    50                      push eax
'0070b1af    8d8d1cfeffff            lea ecx, dword ptr [ebp+fffffe1c]
'0070b1b5    51                      push ecx
'0070b1b6    8d95ecfeffff            lea edx, dword ptr [ebp+fffffeec]
'0070b1bc    52                      push edx
'0070b1bd    ffd6                    call esi
'0070b1bf    50                      push eax
'0070b1c0    8d850cfeffff            lea eax, dword ptr [ebp+fffffe0c]
'0070b1c6    50                      push eax
'0070b1c7    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'0070b1cd    51                      push ecx
'0070b1ce    ffd6                    call esi
'0070b1d0    50                      push eax
'0070b1d1    8d95fcfdffff            lea edx, dword ptr [ebp+fffffdfc]
'0070b1d7    52                      push edx
'0070b1d8    8d85ccfeffff            lea eax, dword ptr [ebp+fffffecc]
'0070b1de    50                      push eax
'0070b1df    ffd6                    call esi
'0070b1e1    50                      push eax
'0070b1e2    8d8decfdffff            lea ecx, dword ptr [ebp+fffffdec]
'0070b1e8    51                      push ecx
'0070b1e9    8d95bcfeffff            lea edx, dword ptr [ebp+fffffebc]
'0070b1ef    52                      push edx
'0070b1f0    ffd6                    call esi
'0070b1f2    50                      push eax
'0070b1f3    8d85dcfdffff            lea eax, dword ptr [ebp+fffffddc]
'0070b1f9    50                      push eax
'0070b1fa    8d8dacfeffff            lea ecx, dword ptr [ebp+fffffeac]
'0070b200    51                      push ecx
'0070b201    ffd6                    call esi
'0070b203    50                      push eax

' *** Reference to "__vbaStrVarMove"
'0070b204    ff153c104000            call dword ptr [0040103c]
'0070b20a    8985a4feffff            mov dword ptr [ebp+fffffea4], eax
'0070b210    899d9cfeffff            mov dword ptr [ebp+fffffe9c], ebx
'0070b216    83ec10                  sub esp, 10
'0070b219    8bd4                    mov edx, esp
'0070b21b    891a                    mov dword ptr [edx], ebx
'0070b21d    8b8da0feffff            mov ecx, dword ptr [ebp+fffffea0]
'0070b223    894a04                  mov dword ptr [edx+04], ecx
'0070b226    894208                  mov dword ptr [edx+08], eax
'0070b229    8b85a8feffff            mov eax, dword ptr [ebp+fffffea8]
'0070b22f    89420c                  mov dword ptr [edx+0c], eax
'0070b232    6a01                    push 01
'0070b234    6880000000              push 00000080
'0070b239    8b7d08                  mov edi, dword ptr [ebp+08]
'0070b23c    8b0f                    mov ecx, dword ptr [edi]
'0070b23e    57                      push edi
'0070b23f    ff9120030000            call dword ptr [ecx+00000320]
'0070b245    50                      push eax
'0070b246    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0070b24c    52                      push edx

' *** Reference to "__vbaObjSet"
'0070b24d    ff15b4104000            call dword ptr [004010b4]
Set var_19 = Nothing
'0070b253    50                      push eax

' *** Reference to "__vbaLateIdCall"
'0070b254    ff1538104000            call dword ptr [00401038]
Call -256 - 20.Member_0x80(#NOT SUPPORTED#)
'0070b25a    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'0070b260    50                      push eax
'0070b261    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'0070b267    51                      push ecx
'0070b268    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0070b26e    52                      push edx
'0070b26f    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0070b271    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_91, var_19)
'0070b277    8d859cfeffff            lea eax, dword ptr [ebp+fffffe9c]
'0070b27d    50                      push eax
'0070b27e    8d8dacfeffff            lea ecx, dword ptr [ebp+fffffeac]
'0070b284    51                      push ecx
'0070b285    8d95bcfeffff            lea edx, dword ptr [ebp+fffffebc]
'0070b28b    52                      push edx
'0070b28c    8d85ccfeffff            lea eax, dword ptr [ebp+fffffecc]
'0070b292    50                      push eax
'0070b293    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'0070b299    51                      push ecx
'0070b29a    8d95ecfeffff            lea edx, dword ptr [ebp+fffffeec]
'0070b2a0    52                      push edx
'0070b2a1    8d85fcfeffff            lea eax, dword ptr [ebp+fffffefc]
'0070b2a7    50                      push eax
'0070b2a8    8d8d0cffffff            lea ecx, dword ptr [ebp+ffffff0c]
'0070b2ae    51                      push ecx
'0070b2af    8d951cffffff            lea edx, dword ptr [ebp+ffffff1c]
'0070b2b5    52                      push edx
'0070b2b6    8d852cffffff            lea eax, dword ptr [ebp+ffffff2c]
'0070b2bc    50                      push eax
'0070b2bd    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'0070b2c3    51                      push ecx
'0070b2c4    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'0070b2ca    52                      push edx
'0070b2cb    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070b2d1    50                      push eax
'0070b2d2    6a0d                    push 0d

' *** Reference to "__vbaFreeVarList"
'0070b2d4    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_88, var_89, var_90, var_93, var_95, var_116, var_117, var_118, var_10, var_115, var_482, var_520, var_822)
'0070b2da    83c464                  add esp, 64
'0070b2dd    b812000000              mov eax, 00000012
'0070b2e2    898594feffff            mov dword ptr [ebp+fffffe94], eax
'0070b2e8    b903000000              mov ecx, 00000003
'0070b2ed    898d8cfeffff            mov dword ptr [ebp+fffffe8c], ecx
'0070b2f3    8b5de0                  mov ebx, dword ptr [ebp-20]
'0070b2f6    66899d74feffff          mov word ptr [ebp+fffffe74], bx
'0070b2fd    ba02000000              mov edx, 00000002
'0070b302    89956cfeffff            mov dword ptr [ebp+fffffe6c], edx
'0070b308    898d54feffff            mov dword ptr [ebp+fffffe54], ecx
'0070b30e    89954cfeffff            mov dword ptr [ebp+fffffe4c], edx
'0070b314    898534feffff            mov dword ptr [ebp+fffffe34], eax
'0070b31a    898d2cfeffff            mov dword ptr [ebp+fffffe2c], ecx
'0070b320    66899d14feffff          mov word ptr [ebp+fffffe14], bx
'0070b327    89950cfeffff            mov dword ptr [ebp+fffffe0c], edx
'0070b32d    c785f4fdffff04000000    mov dword ptr [ebp+fffffdf4], 00000004
'0070b337    8995ecfdffff            mov dword ptr [ebp+fffffdec], edx
'0070b33d    83ec10                  sub esp, 10
'0070b340    8bd4                    mov edx, esp
'0070b342    890a                    mov dword ptr [edx], ecx
'0070b344    8b8d90feffff            mov ecx, dword ptr [ebp+fffffe90]
'0070b34a    894a04                  mov dword ptr [edx+04], ecx
'0070b34d    894208                  mov dword ptr [edx+08], eax
'0070b350    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0070b356    89420c                  mov dword ptr [edx+0c], eax
'0070b359    83ec10                  sub esp, 10
'0070b35c    8bcc                    mov ecx, esp
'0070b35e    8b956cfeffff            mov edx, dword ptr [ebp+fffffe6c]
'0070b364    8911                    mov dword ptr [ecx], edx
'0070b366    8b8570feffff            mov eax, dword ptr [ebp+fffffe70]
'0070b36c    894104                  mov dword ptr [ecx+04], eax
'0070b36f    8b9574feffff            mov edx, dword ptr [ebp+fffffe74]
'0070b375    895108                  mov dword ptr [ecx+08], edx
'0070b378    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'0070b37e    89410c                  mov dword ptr [ecx+0c], eax
'0070b381    83ec10                  sub esp, 10
'0070b384    8bcc                    mov ecx, esp
'0070b386    8b954cfeffff            mov edx, dword ptr [ebp+fffffe4c]
'0070b38c    8911                    mov dword ptr [ecx], edx
'0070b38e    8b8550feffff            mov eax, dword ptr [ebp+fffffe50]
'0070b394    894104                  mov dword ptr [ecx+04], eax
'0070b397    8b9554feffff            mov edx, dword ptr [ebp+fffffe54]
'0070b39d    895108                  mov dword ptr [ecx+08], edx
'0070b3a0    8b8558feffff            mov eax, dword ptr [ebp+fffffe58]
'0070b3a6    89410c                  mov dword ptr [ecx+0c], eax
'0070b3a9    6a03                    push 03
'0070b3ab    689e000000              push 0000009e
'0070b3b0    8b0f                    mov ecx, dword ptr [edi]
'0070b3b2    57                      push edi
'0070b3b3    ff9120030000            call dword ptr [ecx+00000320]
'0070b3b9    50                      push eax
'0070b3ba    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0070b3c0    52                      push edx

' *** Reference to "__vbaObjSet"
'0070b3c1    ff15b4104000            call dword ptr [004010b4]
Set var_87 = Nothing
'0070b3c7    50                      push eax
'0070b3c8    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070b3ce    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'0070b3cf    ff157c114000            call dword ptr [0040117c]
var_88 = var_87.UNK_-256 - 20_158
'0070b3d5    83c440                  add esp, 40
'0070b3d8    50                      push eax
'0070b3d9    83ec10                  sub esp, 10
'0070b3dc    8bcc                    mov ecx, esp
'0070b3de    8b952cfeffff            mov edx, dword ptr [ebp+fffffe2c]
'0070b3e4    8911                    mov dword ptr [ecx], edx
'0070b3e6    8b8530feffff            mov eax, dword ptr [ebp+fffffe30]
'0070b3ec    894104                  mov dword ptr [ecx+04], eax
'0070b3ef    8b9534feffff            mov edx, dword ptr [ebp+fffffe34]
'0070b3f5    895108                  mov dword ptr [ecx+08], edx
'0070b3f8    8b8538feffff            mov eax, dword ptr [ebp+fffffe38]
'0070b3fe    89410c                  mov dword ptr [ecx+0c], eax
'0070b401    83ec10                  sub esp, 10
'0070b404    8bcc                    mov ecx, esp
'0070b406    8b950cfeffff            mov edx, dword ptr [ebp+fffffe0c]
'0070b40c    8911                    mov dword ptr [ecx], edx
'0070b40e    8b8510feffff            mov eax, dword ptr [ebp+fffffe10]
'0070b414    894104                  mov dword ptr [ecx+04], eax
'0070b417    8b9514feffff            mov edx, dword ptr [ebp+fffffe14]
'0070b41d    895108                  mov dword ptr [ecx+08], edx
'0070b420    8b8518feffff            mov eax, dword ptr [ebp+fffffe18]
'0070b426    89410c                  mov dword ptr [ecx+0c], eax
'0070b429    83ec10                  sub esp, 10
'0070b42c    8bcc                    mov ecx, esp
'0070b42e    8b95ecfdffff            mov edx, dword ptr [ebp+fffffdec]
'0070b434    8911                    mov dword ptr [ecx], edx
'0070b436    8b85f0fdffff            mov eax, dword ptr [ebp+fffffdf0]
'0070b43c    894104                  mov dword ptr [ecx+04], eax
'0070b43f    8b95f4fdffff            mov edx, dword ptr [ebp+fffffdf4]
'0070b445    895108                  mov dword ptr [ecx+08], edx
'0070b448    8b85f8fdffff            mov eax, dword ptr [ebp+fffffdf8]
'0070b44e    89410c                  mov dword ptr [ecx+0c], eax
'0070b451    6a03                    push 03
'0070b453    689e000000              push 0000009e
'0070b458    8b0f                    mov ecx, dword ptr [edi]
'0070b45a    57                      push edi
'0070b45b    ff9120030000            call dword ptr [ecx+00000320]
'0070b461    50                      push eax
'0070b462    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'0070b468    52                      push edx

' *** Reference to "__vbaObjSet"
'0070b469    ff15b4104000            call dword ptr [004010b4]
Set var_91 = Nothing
'0070b46f    50                      push eax
'0070b470    8d854cffffff            lea eax, dword ptr [ebp+ffffff4c]
'0070b476    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'0070b477    ff157c114000            call dword ptr [0040117c]
var_89 = var_91.UNK_-256 - 20_158
'0070b47d    83c440                  add esp, 40
'0070b480    50                      push eax

' *** Reference to "__vbaVarTstGe"
'0070b481    ff15a8124000            call dword ptr [004012a8]
'0070b487    66898590fdffff          mov word ptr [ebp+fffffd90], ax
'0070b48e    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'0070b494    51                      push ecx
'0070b495    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0070b49b    52                      push edx
'0070b49c    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070b49e    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_91)
'0070b4a4    8d854cffffff            lea eax, dword ptr [ebp+ffffff4c]
'0070b4aa    50                      push eax
'0070b4ab    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]
'0070b4b1    51                      push ecx
'0070b4b2    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'0070b4b4    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_88, var_89)
'0070b4ba    83c418                  add esp, 18
'0070b4bd    6683bd90fdffff00        cmp word ptr [ebp+fffffd90], 00
'0070b4c5    0f847b030000            je 70b846

If (((var_87) >= (var_89))) Then
'0070b4cb    8b17                    mov edx, dword ptr [edi]
'0070b4cd    57                      push edi
'0070b4ce    ff920c030000            call dword ptr [edx+0000030c]
'0070b4d4    50                      push eax
'0070b4d5    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'0070b4db    50                      push eax

' *** Reference to "__vbaObjSet"
'0070b4dc    ff15b4104000            call dword ptr [004010b4]
    Set var_94 = 
'0070b4e2    898588fdffff            mov dword ptr [ebp+fffffd88], eax
'0070b4e8    8b0f                    mov ecx, dword ptr [edi]
'0070b4ea    57                      push edi
'0070b4eb    ff910c030000            call dword ptr [ecx+0000030c]
'0070b4f1    50                      push eax
'0070b4f2    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'0070b4f8    52                      push edx

' *** Reference to "__vbaObjSet"
'0070b4f9    ff15b4104000            call dword ptr [004010b4]
    Set var_91 = var_94
'0070b4ff    898590fdffff            mov dword ptr [ebp+fffffd90], eax
'0070b505    8b08                    mov ecx, dword ptr [eax]
'0070b507    8d5594                  lea edx, dword ptr [ebp-6c]
'0070b50a    52                      push edx
'0070b50b    50                      push eax
'0070b50c    ff91a0000000            call dword ptr [ecx+000000a0]
'0070b512    dbe2                    fnclex
'0070b514    85c0                    test ax, ax
'0070b516    7d18                    jge 70b530
    
    If (    var_91) Then
'0070b518    68a0000000              push 000000a0
'0070b51d    68d00d4300              push 00430dd0
'0070b522    8b8d90fdffff            mov ecx, dword ptr [ebp+fffffd90]
'0070b528    51                      push ecx
'0070b529    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070b52a    ff1580104000            call dword ptr [00401080]
    
End If
'0070b530    8b4594                  mov eax, dword ptr [ebp-6c]
'0070b533    c7459400000000          mov dword ptr [ebp-6c], 00000000
'0070b53a    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'0070b540    ba08000000              mov edx, 00000008
'0070b545    89955cffffff            mov dword ptr [ebp+ffffff5c], edx
'0070b54b    33c0                    xor eax, eax
var_num1 = Empty
'0070b54d    898594feffff            mov dword ptr [ebp+fffffe94], eax
'0070b553    b903000000              mov ecx, 00000003
'0070b558    898d8cfeffff            mov dword ptr [ebp+fffffe8c], ecx
'0070b55e    66899d74feffff          mov word ptr [ebp+fffffe74], bx
'0070b565    c7856cfeffff02000000    mov dword ptr [ebp+fffffe6c], 00000002
'0070b56f    898554feffff            mov dword ptr [ebp+fffffe54], eax
'0070b575    c7854cfeffff02000000    mov dword ptr [ebp+fffffe4c], 00000002
'0070b57f    c78534fefffff8624500    mov dword ptr [ebp+fffffe34], 004562f8
'0070b589    89952cfeffff            mov dword ptr [ebp+fffffe2c], edx
'0070b58f    898524feffff            mov dword ptr [ebp+fffffe24], eax
'0070b595    898d1cfeffff            mov dword ptr [ebp+fffffe1c], ecx
'0070b59b    66899d04feffff          mov word ptr [ebp+fffffe04], bx
'0070b5a2    bb02000000              mov ebx, 00000002
'0070b5a7    899dfcfdffff            mov dword ptr [ebp+fffffdfc], ebx
'0070b5ad    c785e4fdffff01000000    mov dword ptr [ebp+fffffde4], 00000001
'0070b5b7    899ddcfdffff            mov dword ptr [ebp+fffffddc], ebx
'0070b5bd    c785c4fdffff01000000    mov dword ptr [ebp+fffffdc4], 00000001
'0070b5c7    899dbcfdffff            mov dword ptr [ebp+fffffdbc], ebx
'0070b5cd    c785b4fdffff30634500    mov dword ptr [ebp+fffffdb4], 00456330
'0070b5d7    8995acfdffff            mov dword ptr [ebp+fffffdac], edx
'0070b5dd    c785a4fdffff70084300    mov dword ptr [ebp+fffffda4], 00430870
'0070b5e7    89959cfdffff            mov dword ptr [ebp+fffffd9c], edx
'0070b5ed    8b9588fdffff            mov edx, dword ptr [ebp+fffffd88]
'0070b5f3    8b1a                    mov ebx, dword ptr [edx]
'0070b5f5    8d955cffffff            lea edx, dword ptr [ebp+ffffff5c]
'0070b5fb    52                      push edx
'0070b5fc    83ec10                  sub esp, 10
'0070b5ff    8bd4                    mov edx, esp
'0070b601    890a                    mov dword ptr [edx], ecx
'0070b603    8b8d90feffff            mov ecx, dword ptr [ebp+fffffe90]
'0070b609    894a04                  mov dword ptr [edx+04], ecx
'0070b60c    894208                  mov dword ptr [edx+08], eax
'0070b60f    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0070b615    89420c                  mov dword ptr [edx+0c], eax
'0070b618    83ec10                  sub esp, 10
'0070b61b    8bcc                    mov ecx, esp
'0070b61d    8b956cfeffff            mov edx, dword ptr [ebp+fffffe6c]
'0070b623    8911                    mov dword ptr [ecx], edx
'0070b625    8b8570feffff            mov eax, dword ptr [ebp+fffffe70]
'0070b62b    894104                  mov dword ptr [ecx+04], eax
'0070b62e    8b9574feffff            mov edx, dword ptr [ebp+fffffe74]
'0070b634    895108                  mov dword ptr [ecx+08], edx
'0070b637    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'0070b63d    89410c                  mov dword ptr [ecx+0c], eax
'0070b640    83ec10                  sub esp, 10
'0070b643    8bcc                    mov ecx, esp
'0070b645    8b954cfeffff            mov edx, dword ptr [ebp+fffffe4c]
'0070b64b    8911                    mov dword ptr [ecx], edx
'0070b64d    8b8550feffff            mov eax, dword ptr [ebp+fffffe50]
'0070b653    894104                  mov dword ptr [ecx+04], eax
'0070b656    8b9554feffff            mov edx, dword ptr [ebp+fffffe54]
'0070b65c    895108                  mov dword ptr [ecx+08], edx
'0070b65f    8b8558feffff            mov eax, dword ptr [ebp+fffffe58]
'0070b665    89410c                  mov dword ptr [ecx+0c], eax
'0070b668    6a03                    push 03
'0070b66a    689e000000              push 0000009e
'0070b66f    8b0f                    mov ecx, dword ptr [edi]
'0070b671    57                      push edi
'0070b672    ff9120030000            call dword ptr [ecx+00000320]
'0070b678    50                      push eax
'0070b679    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0070b67f    52                      push edx

' *** Reference to "__vbaObjSet"
'0070b680    ff15b4104000            call dword ptr [004010b4]
Set var_87 = Nothing
'0070b686    50                      push eax
'0070b687    8d854cffffff            lea eax, dword ptr [ebp+ffffff4c]
'0070b68d    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'0070b68e    ff157c114000            call dword ptr [0040117c]
var_89 = var_87.UNK_-256 - 20_158
'0070b694    83c440                  add esp, 40
'0070b697    50                      push eax
'0070b698    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'0070b69e    51                      push ecx
'0070b69f    ffd6                    call esi
'0070b6a1    50                      push eax
'0070b6a2    8d952cfeffff            lea edx, dword ptr [ebp+fffffe2c]
'0070b6a8    52                      push edx
'0070b6a9    8d852cffffff            lea eax, dword ptr [ebp+ffffff2c]
'0070b6af    50                      push eax
'0070b6b0    ffd6                    call esi
'0070b6b2    50                      push eax
'0070b6b3    83ec10                  sub esp, 10
'0070b6b6    8bcc                    mov ecx, esp
'0070b6b8    8b951cfeffff            mov edx, dword ptr [ebp+fffffe1c]
'0070b6be    8911                    mov dword ptr [ecx], edx
'0070b6c0    8b8520feffff            mov eax, dword ptr [ebp+fffffe20]
'0070b6c6    894104                  mov dword ptr [ecx+04], eax
'0070b6c9    8b9524feffff            mov edx, dword ptr [ebp+fffffe24]
'0070b6cf    895108                  mov dword ptr [ecx+08], edx
'0070b6d2    8b8528feffff            mov eax, dword ptr [ebp+fffffe28]
'0070b6d8    89410c                  mov dword ptr [ecx+0c], eax
'0070b6db    83ec10                  sub esp, 10
'0070b6de    8bcc                    mov ecx, esp
'0070b6e0    8b95fcfdffff            mov edx, dword ptr [ebp+fffffdfc]
'0070b6e6    8911                    mov dword ptr [ecx], edx
'0070b6e8    8b8500feffff            mov eax, dword ptr [ebp+fffffe00]
'0070b6ee    894104                  mov dword ptr [ecx+04], eax
'0070b6f1    8b9504feffff            mov edx, dword ptr [ebp+fffffe04]
'0070b6f7    895108                  mov dword ptr [ecx+08], edx
'0070b6fa    8b8508feffff            mov eax, dword ptr [ebp+fffffe08]
'0070b700    89410c                  mov dword ptr [ecx+0c], eax
'0070b703    83ec10                  sub esp, 10
'0070b706    8bcc                    mov ecx, esp
'0070b708    8b95dcfdffff            mov edx, dword ptr [ebp+fffffddc]
'0070b70e    8911                    mov dword ptr [ecx], edx
'0070b710    8b85e0fdffff            mov eax, dword ptr [ebp+fffffde0]
'0070b716    894104                  mov dword ptr [ecx+04], eax
'0070b719    8b95e4fdffff            mov edx, dword ptr [ebp+fffffde4]
'0070b71f    895108                  mov dword ptr [ecx+08], edx
'0070b722    8b85e8fdffff            mov eax, dword ptr [ebp+fffffde8]
'0070b728    89410c                  mov dword ptr [ecx+0c], eax
'0070b72b    6a03                    push 03
'0070b72d    689e000000              push 0000009e
'0070b732    8b0f                    mov ecx, dword ptr [edi]
'0070b734    57                      push edi
'0070b735    ff9120030000            call dword ptr [ecx+00000320]
'0070b73b    50                      push eax
'0070b73c    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0070b742    52                      push edx

' *** Reference to "__vbaObjSet"
'0070b743    ff15b4104000            call dword ptr [004010b4]
Set var_19 = Nothing
'0070b749    50                      push eax
'0070b74a    8d851cffffff            lea eax, dword ptr [ebp+ffffff1c]
'0070b750    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'0070b751    ff157c114000            call dword ptr [0040117c]
var_95 = var_19.UNK_-256 - 20_158
'0070b757    83c440                  add esp, 40
'0070b75a    50                      push eax
'0070b75b    8d8dbcfdffff            lea ecx, dword ptr [ebp+fffffdbc]
'0070b761    51                      push ecx
'0070b762    8d950cffffff            lea edx, dword ptr [ebp+ffffff0c]
'0070b768    52                      push edx

' *** Reference to "__vbaVarAdd"
'0070b769    ff158c124000            call dword ptr [0040128c]
'0070b76f    50                      push eax
'0070b770    8d85fcfeffff            lea eax, dword ptr [ebp+fffffefc]
'0070b776    50                      push eax
'0070b777    ffd6                    call esi
'0070b779    50                      push eax
'0070b77a    8d8dacfdffff            lea ecx, dword ptr [ebp+fffffdac]
'0070b780    51                      push ecx
'0070b781    8d95ecfeffff            lea edx, dword ptr [ebp+fffffeec]
'0070b787    52                      push edx
'0070b788    ffd6                    call esi
'0070b78a    50                      push eax
'0070b78b    8d859cfdffff            lea eax, dword ptr [ebp+fffffd9c]
'0070b791    50                      push eax
'0070b792    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'0070b798    51                      push ecx
'0070b799    ffd6                    call esi
'0070b79b    50                      push eax
'0070b79c    8d5590                  lea edx, dword ptr [ebp-70]
'0070b79f    52                      push edx

' *** Reference to "__vbaStrVarVal"
'0070b7a0    ff15fc114000            call dword ptr [004011fc]
'0070b7a6    50                      push eax
'0070b7a7    8bb588fdffff            mov esi, dword ptr [ebp+fffffd88]
'0070b7ad    56                      push esi
'0070b7ae    ff93a4000000            call dword ptr [ebx+000000a4]
'0070b7b4    dbe2                    fnclex
'0070b7b6    85c0                    test ax, ax
'0070b7b8    7d12                    jge 70b7cc
'0070b7ba    68a4000000              push 000000a4
'0070b7bf    68d00d4300              push 00430dd0
'0070b7c4    56                      push esi
'0070b7c5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070b7c6    ff1580104000            call dword ptr [00401080]
'0070b7cc    8d4d90                  lea ecx, dword ptr [ebp-70]

' *** Reference to "__vbaFreeStr"
'0070b7cf    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_8)
'0070b7d5    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'0070b7db    50                      push eax
'0070b7dc    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'0070b7e2    51                      push ecx
'0070b7e3    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'0070b7e9    52                      push edx
'0070b7ea    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0070b7f0    50                      push eax
'0070b7f1    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'0070b7f3    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_91, var_19, var_94)
'0070b7f9    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'0070b7ff    51                      push ecx
'0070b800    8d95ecfeffff            lea edx, dword ptr [ebp+fffffeec]
'0070b806    52                      push edx
'0070b807    8d85fcfeffff            lea eax, dword ptr [ebp+fffffefc]
'0070b80d    50                      push eax
'0070b80e    8d8d0cffffff            lea ecx, dword ptr [ebp+ffffff0c]
'0070b814    51                      push ecx
'0070b815    8d952cffffff            lea edx, dword ptr [ebp+ffffff2c]
'0070b81b    52                      push edx
'0070b81c    8d851cffffff            lea eax, dword ptr [ebp+ffffff1c]
'0070b822    50                      push eax
'0070b823    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'0070b829    51                      push ecx
'0070b82a    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'0070b830    52                      push edx
'0070b831    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070b837    50                      push eax
'0070b838    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0070b83a    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_88, var_89, var_90, var_95, var_93, var_116, var_117, var_118, var_10)
'0070b840    83c43c                  add esp, 3c
'0070b843    8b5de0                  mov ebx, dword ptr [ebp-20]

'ERROR: Two many next close:
End If
'0070b846    8b45a4                  mov eax, dword ptr [ebp-5c]
'0070b849    8b08                    mov ecx, dword ptr [eax]
'0070b84b    50                      push eax
'0070b84c    ff91ec000000            call dword ptr [ecx+000000ec]
'0070b852    dbe2                    fnclex
'0070b854    85c0                    test ax, ax
'0070b856    7d15                    jge 70b86d
'0070b858    68ec000000              push 000000ec
'0070b85d    6830314300              push 00433130
'0070b862    8b55a4                  mov edx, dword ptr [ebp-5c]
'0070b865    52                      push edx
'0070b866    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070b867    ff1580104000            call dword ptr [00401080]
'0070b86d    8b45d8                  mov eax, dword ptr [ebp-28]
'0070b870    8b08                    mov ecx, dword ptr [eax]
'0070b872    50                      push eax
'0070b873    ff91c4000000            call dword ptr [ecx+000000c4]
'0070b879    dbe2                    fnclex
'0070b87b    85c0                    test ax, ax
'0070b87d    7d15                    jge 70b894
'0070b87f    68c4000000              push 000000c4
'0070b884    6830314300              push 00433130
'0070b889    8b55d8                  mov edx, dword ptr [ebp-28]
'0070b88c    52                      push edx
'0070b88d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070b88e    ff1580104000            call dword ptr [00401080]
'0070b894    6683c301                add bx, 01
var_num2 = var_3 + 1
'0070b898    0f80a9020000            jo 70bb47
'0070b89e    895de0                  mov dword ptr [ebp-20], ebx

' *** Reference to "__vbaObjSet"
'0070b8a1    8b35b4104000            mov esi, dword ptr [004010b4]
'0070b8a7    33db                    xor ebx, ebx
var_num2 = Empty
'0070b8a9    e9a7ceffff              jmp 708755

'ERROR: Two many next close:
Loop
'0070b8ae    8b08                    mov ecx, dword ptr [eax]
'0070b8b0    50                      push eax
'0070b8b1    ff91c4000000            call dword ptr [ecx+000000c4]
'0070b8b7    dbe2                    fnclex
'0070b8b9    3bc3                    cmp eax, ebx
'0070b8bb    7d15                    jge 70b8d2

If (var_83 < 0) Then
'0070b8bd    68c4000000              push 000000c4
'0070b8c2    6830314300              push 00433130
'0070b8c7    8b55a4                  mov edx, dword ptr [ebp-5c]
'0070b8ca    52                      push edx
'0070b8cb    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070b8cc    ff1580104000            call dword ptr [00401080]
    
End If
'0070b8d2    8b45c4                  mov eax, dword ptr [ebp-3c]
'0070b8d5    8b08                    mov ecx, dword ptr [eax]
'0070b8d7    50                      push eax
'0070b8d8    ff91c4000000            call dword ptr [ecx+000000c4]
'0070b8de    dbe2                    fnclex
'0070b8e0    3bc3                    cmp eax, ebx
'0070b8e2    7d15                    jge 70b8f9

If (var_9 < 0) Then
'0070b8e4    68c4000000              push 000000c4
'0070b8e9    6830314300              push 00433130
'0070b8ee    8b55c4                  mov edx, dword ptr [ebp-3c]
'0070b8f1    52                      push edx
'0070b8f2    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070b8f3    ff1580104000            call dword ptr [00401080]
    
End If
'0070b8f9    8b45b8                  mov eax, dword ptr [ebp-48]
'0070b8fc    8b08                    mov ecx, dword ptr [eax]
'0070b8fe    50                      push eax
'0070b8ff    ff91c4000000            call dword ptr [ecx+000000c4]
'0070b905    dbe2                    fnclex
'0070b907    3bc3                    cmp eax, ebx
'0070b909    7d15                    jge 70b920

If (var_61 < 0) Then
'0070b90b    68c4000000              push 000000c4
'0070b910    6830314300              push 00433130
'0070b915    8b55b8                  mov edx, dword ptr [ebp-48]
'0070b918    52                      push edx
'0070b919    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070b91a    ff1580104000            call dword ptr [00401080]
    
End If
'0070b920    b801000000              mov eax, 00000001
'0070b925    898594feffff            mov dword ptr [ebp+fffffe94], eax
'0070b92b    b903000000              mov ecx, 00000003
'0070b930    898d8cfeffff            mov dword ptr [ebp+fffffe8c], ecx
'0070b936    83ec10                  sub esp, 10
'0070b939    8bd4                    mov edx, esp
'0070b93b    890a                    mov dword ptr [edx], ecx
'0070b93d    8b8d90feffff            mov ecx, dword ptr [ebp+fffffe90]
'0070b943    894a04                  mov dword ptr [edx+04], ecx
'0070b946    894208                  mov dword ptr [edx+08], eax
'0070b949    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0070b94f    89420c                  mov dword ptr [edx+0c], eax
'0070b952    6a11                    push 11
'0070b954    8b0f                    mov ecx, dword ptr [edi]
'0070b956    57                      push edi
'0070b957    ff9120030000            call dword ptr [ecx+00000320]
'0070b95d    50                      push eax
'0070b95e    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0070b964    52                      push edx
'0070b965    ffd6                    call esi
Set var_87 = Nothing
'0070b967    50                      push eax

' *** Reference to "__vbaLateIdSt"
'0070b968    ff15ec124000            call dword ptr [004012ec]
var_87.[UNMANAGED] = 1
'0070b96e    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeObj"
'0070b974    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_87)
'0070b97a    b805000000              mov eax, 00000005
'0070b97f    898594feffff            mov dword ptr [ebp+fffffe94], eax
'0070b985    b903000000              mov ecx, 00000003
'0070b98a    898d8cfeffff            mov dword ptr [ebp+fffffe8c], ecx
'0070b990    83ec10                  sub esp, 10
'0070b993    8bd4                    mov edx, esp
'0070b995    890a                    mov dword ptr [edx], ecx
'0070b997    8b8d90feffff            mov ecx, dword ptr [ebp+fffffe90]
'0070b99d    894a04                  mov dword ptr [edx+04], ecx
'0070b9a0    894208                  mov dword ptr [edx+08], eax
'0070b9a3    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0070b9a9    89420c                  mov dword ptr [edx+0c], eax
'0070b9ac    6a12                    push 12
'0070b9ae    8b0f                    mov ecx, dword ptr [edi]
'0070b9b0    57                      push edi
'0070b9b1    ff9120030000            call dword ptr [ecx+00000320]
'0070b9b7    50                      push eax
'0070b9b8    8d9578ffffff            lea edx, dword ptr [ebp+ffffff78]
'0070b9be    52                      push edx
'0070b9bf    ffd6                    call esi
Set var_87 = Nothing
'0070b9c1    50                      push eax

' *** Reference to "__vbaLateIdSt"
'0070b9c2    ff15ec124000            call dword ptr [004012ec]
var_87.[UNMANAGED] = 5
'0070b9c8    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeObj"
'0070b9ce    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_87)
'0070b9d4    391d24be7200            cmp dword ptr [0072be24], ebx
'0070b9da    7510                    jne 70b9ec
'0070b9dc    6824be7200              push 0072be24
'0070b9e1    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0070b9e6    ff152c124000            call dword ptr [0040122c]
Dim var_2 As New Global
'0070b9ec    8b3524be7200            mov esi, dword ptr [0072be24]
'0070b9f2    8b06                    mov eax, dword ptr [esi]
'0070b9f4    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]
'0070b9fa    51                      push ecx
'0070b9fb    56                      push esi
'0070b9fc    ff5018                  call dword ptr [eax+18]
Set var_87 = var_2.Screen
'0070b9ff    dbe2                    fnclex
'0070ba01    3bc3                    cmp eax, ebx
'0070ba03    7d0f                    jge 70ba14

If (var_2.Screen < 0) Then
'0070ba05    6a18                    push 18
'0070ba07    6860004300              push 00430060
'0070ba0c    56                      push esi
'0070ba0d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070ba0e    ff1580104000            call dword ptr [00401080]
    
End If
'0070ba14    8bb578ffffff            mov esi, dword ptr [ebp+ffffff78]
'0070ba1a    8b3e                    mov edi, dword ptr [esi]
'0070ba1c    33c9                    xor ecx, ecx

' *** Reference to "__vbaI2I4"
'0070ba1e    ff1550114000            call dword ptr [00401150]
'0070ba24    50                      push eax
'0070ba25    56                      push esi
'0070ba26    ff577c                  call dword ptr [edi+7c]
var_87.MousePointer = CInt(0)
'0070ba29    dbe2                    fnclex
'0070ba2b    3bc3                    cmp eax, ebx
'0070ba2d    7d0f                    jge 70ba3e
'0070ba2f    6a7c                    push 7c
'0070ba31    6810be4300              push 0043be10
'0070ba36    56                      push esi
'0070ba37    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070ba38    ff1580104000            call dword ptr [00401080]
'0070ba3e    8d8d78ffffff            lea ecx, dword ptr [ebp+ffffff78]

' *** Reference to "__vbaFreeObj"
'0070ba44    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_87)

' *** Reference to "__vbaExitProc"
'0070ba4a    ff15a0104000            call dword ptr [004010a0]
'0070ba50    9b                      fwait
'0070ba51    6828bb7000              push 0070bb28
'0070ba56    e9b2000000              jmp 70bb0d
'0070ba5b    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'0070ba61    52                      push edx
'0070ba62    8d4580                  lea eax, dword ptr [ebp-80]
'0070ba65    50                      push eax
'0070ba66    8d4d84                  lea ecx, dword ptr [ebp-7c]
'0070ba69    51                      push ecx
'0070ba6a    8d5588                  lea edx, dword ptr [ebp-78]
'0070ba6d    52                      push edx
'0070ba6e    8d458c                  lea eax, dword ptr [ebp-74]
'0070ba71    50                      push eax
'0070ba72    8d4d90                  lea ecx, dword ptr [ebp-70]
'0070ba75    51                      push ecx
'0070ba76    8d5594                  lea edx, dword ptr [ebp-6c]
'0070ba79    52                      push edx
'0070ba7a    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0070ba7c    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_121, , -4540, -4544, -4548, -4552, -4556)
'0070ba82    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'0070ba88    50                      push eax
'0070ba89    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'0070ba8f    51                      push ecx
'0070ba90    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'0070ba96    52                      push edx
'0070ba97    8d8578ffffff            lea eax, dword ptr [ebp+ffffff78]
'0070ba9d    50                      push eax
'0070ba9e    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'0070baa0    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_87, var_91, var_19, var_94)
'0070baa6    8d8d9cfeffff            lea ecx, dword ptr [ebp+fffffe9c]
'0070baac    51                      push ecx
'0070baad    8d95acfeffff            lea edx, dword ptr [ebp+fffffeac]
'0070bab3    52                      push edx
'0070bab4    8d85bcfeffff            lea eax, dword ptr [ebp+fffffebc]
'0070baba    50                      push eax
'0070babb    8d8dccfeffff            lea ecx, dword ptr [ebp+fffffecc]
'0070bac1    51                      push ecx
'0070bac2    8d95dcfeffff            lea edx, dword ptr [ebp+fffffedc]
'0070bac8    52                      push edx
'0070bac9    8d85ecfeffff            lea eax, dword ptr [ebp+fffffeec]
'0070bacf    50                      push eax
'0070bad0    8d8dfcfeffff            lea ecx, dword ptr [ebp+fffffefc]
'0070bad6    51                      push ecx
'0070bad7    8d950cffffff            lea edx, dword ptr [ebp+ffffff0c]
'0070badd    52                      push edx
'0070bade    8d851cffffff            lea eax, dword ptr [ebp+ffffff1c]
'0070bae4    50                      push eax
'0070bae5    8d8d2cffffff            lea ecx, dword ptr [ebp+ffffff2c]
'0070baeb    51                      push ecx
'0070baec    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'0070baf2    52                      push edx
'0070baf3    8d854cffffff            lea eax, dword ptr [ebp+ffffff4c]
'0070baf9    50                      push eax
'0070bafa    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]
'0070bb00    51                      push ecx
'0070bb01    6a0d                    push 0d

' *** Reference to "__vbaFreeVarList"
'0070bb03    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_88, var_89, var_90, var_93, var_95, var_116, var_117, var_118, var_10, var_115, var_482, var_520, var_822)
'0070bb09    83c46c                  add esp, 6c
'0070bb0c    c3                      ret
'0070bb0d    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'0070bb10    8b3508134000            mov esi, dword ptr [00401308]
'0070bb16    ffd6                    call esi
'#Cleanup(var_45)
'0070bb18    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0070bb1b    ffd6                    call esi
'#Cleanup(var_9)
'0070bb1d    8d4db8                  lea ecx, dword ptr [ebp-48]
'0070bb20    ffd6                    call esi
'#Cleanup(var_61)
'0070bb22    8d4da4                  lea ecx, dword ptr [ebp-5c]
'0070bb25    ffd6                    call esi
'#Cleanup(var_83)
'0070bb27    c3                      ret
'0070bb28    8b4508                  mov eax, dword ptr [ebp+08]
'0070bb2b    8b10                    mov edx, dword ptr [eax]
'0070bb2d    50                      push eax
'0070bb2e    ff5208                  call dword ptr [edx+08]
'0070bb31    8b45f4                  mov eax, dword ptr [ebp-0c]
'0070bb34    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'0070bb37    64890d00000000          mov dword ptr fs:[00000000], ecx
'0070bb3e    5f                      pop edi
'0070bb3f    5e                      pop esi
'0070bb40    5b                      pop ebx
'0070bb41    8be5                    mov esp, ebp
'0070bb43    5d                      pop ebp
'0070bb44    c20400                  ret 0004


End Function


'Event for Form
Private Sub Form_Load()
'0070e3e0    55                      push ebp
'0070e3e1    8bec                    mov ebp, esp
'0070e3e3    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'0070e3e6    6866724000              push 00407266
'0070e3eb    64a100000000            mov ax, word ptr fs:[00000000]
'0070e3f1    50                      push eax
    'Reference to '__except_list'
'0070e3f2    64892500000000          mov dword ptr fs:[00000000], esp
'0070e3f9    81ec2c010000            sub esp, 0000012c
'0070e3ff    53                      push ebx
'0070e400    56                      push esi
'0070e401    57                      push edi
'0070e402    8965ec                  mov dword ptr [ebp-14], esp
'0070e405    c745f0606c4000          mov dword ptr [ebp-10], 00406c60
'0070e40c    8b7d08                  mov edi, dword ptr [ebp+08]
'0070e40f    8bc7                    mov eax, edi
'0070e411    83e001                  and eax, 01
'0070e414    8945f4                  mov dword ptr [ebp-0c], eax
'0070e417    83e7fe                  and edi, fffffffe
'0070e41a    897d08                  mov dword ptr [ebp+08], edi
'0070e41d    33f6                    xor esi, esi
'0070e41f    8975f8                  mov dword ptr [ebp-08], esi
'0070e422    8b0f                    mov ecx, dword ptr [edi]
'0070e424    57                      push edi
'0070e425    ff5104                  call dword ptr [ecx+04]
'0070e428    8975e0                  mov dword ptr [ebp-20], esi
'0070e42b    8975dc                  mov dword ptr [ebp-24], esi
'0070e42e    8975d8                  mov dword ptr [ebp-28], esi
'0070e431    8975d4                  mov dword ptr [ebp-2c], esi
'0070e434    8975d0                  mov dword ptr [ebp-30], esi
'0070e437    8975cc                  mov dword ptr [ebp-34], esi
'0070e43a    8975c8                  mov dword ptr [ebp-38], esi
'0070e43d    8975c4                  mov dword ptr [ebp-3c], esi
'0070e440    8975c0                  mov dword ptr [ebp-40], esi
'0070e443    8975b0                  mov dword ptr [ebp-50], esi
'0070e446    8975a0                  mov dword ptr [ebp-60], esi
'0070e449    897590                  mov dword ptr [ebp-70], esi
'0070e44c    897580                  mov dword ptr [ebp-80], esi
'0070e44f    89b570ffffff            mov dword ptr [ebp+ffffff70], esi
'0070e455    89b560ffffff            mov dword ptr [ebp+ffffff60], esi
'0070e45b    89b550ffffff            mov dword ptr [ebp+ffffff50], esi
'0070e461    89b540ffffff            mov dword ptr [ebp+ffffff40], esi
'0070e467    89b530ffffff            mov dword ptr [ebp+ffffff30], esi
'0070e46d    89b520ffffff            mov dword ptr [ebp+ffffff20], esi
'0070e473    89b510ffffff            mov dword ptr [ebp+ffffff10], esi
'0070e479    89b500ffffff            mov dword ptr [ebp+ffffff00], esi
'0070e47f    89b5f0feffff            mov dword ptr [ebp+fffffef0], esi
'0070e485    89b5e0feffff            mov dword ptr [ebp+fffffee0], esi
'0070e48b    89b5dcfeffff            mov dword ptr [ebp+fffffedc], esi
'0070e491    66393528b07200          cmp word ptr [0072b028], si
'0070e498    7508                    jne 70e4a2
'0070e49a    6a01                    push 01

' *** Reference to "__vbaOnError"
'0070e49c    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'0070e4a2    8b07                    mov eax, dword ptr [edi]
'0070e4a4    57                      push edi
'0070e4a5    ff90f8060000            call dword ptr [eax+000006f8]
'0070e4ab    3bc6                    cmp eax, esi
'0070e4ad    7d12                    jge 70e4c1

If (frmExperience < 0) Then
'0070e4af    68f8060000              push 000006f8
'0070e4b4    68481d4300              push 00431d48
'0070e4b9    57                      push edi
'0070e4ba    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070e4bb    ff1580104000            call dword ptr [00401080]
    
End If

' *** Reference to "__vbaExitProc"
'0070e4c1    ff15a0104000            call dword ptr [004010a0]
'0070e4c7    687ae87000              push 0070e87a
'0070e4cc    e9a8030000              jmp 70e879

' *** Reference to "rtcErrObj"
'0070e4d1    8b1d6c124000            mov ebx, dword ptr [0040126c]
'0070e4d7    ffd3                    call ebx
'0070e4d9    50                      push eax
'0070e4da    8d55c4                  lea edx, dword ptr [ebp-3c]
'0070e4dd    52                      push edx

' *** Reference to "__vbaObjSet"
'0070e4de    8b3db4104000            mov edi, dword ptr [004010b4]
'0070e4e4    ffd7                    call edi
Set var_9 = Err
'0070e4e6    8bf0                    mov esi, eax
'0070e4e8    8b06                    mov eax, dword ptr [esi]
'0070e4ea    8d4de0                  lea ecx, dword ptr [ebp-20]
'0070e4ed    51                      push ecx
'0070e4ee    56                      push esi
'0070e4ef    ff502c                  call dword ptr [eax+2c]
var_3 = var_9.Description
'0070e4f2    dbe2                    fnclex
'0070e4f4    85c0                    test ax, ax
'0070e4f6    7d0f                    jge 70e507
'0070e4f8    6a2c                    push 2c
'0070e4fa    685c084300              push 0043085c
'0070e4ff    56                      push esi
'0070e500    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070e501    ff1580104000            call dword ptr [00401080]
'0070e507    ffd3                    call ebx
'0070e509    50                      push eax
'0070e50a    8d55c0                  lea edx, dword ptr [ebp-40]
'0070e50d    52                      push edx
'0070e50e    ffd7                    call edi
Set var_5 = Err
'0070e510    8bf0                    mov esi, eax
'0070e512    8b06                    mov eax, dword ptr [esi]
'0070e514    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'0070e51a    51                      push ecx
'0070e51b    56                      push esi
'0070e51c    ff501c                  call dword ptr [eax+1c]
var_10 = var_5.Number
'0070e51f    dbe2                    fnclex
'0070e521    85c0                    test ax, ax
'0070e523    7d0f                    jge 70e534
'0070e525    6a1c                    push 1c
'0070e527    685c084300              push 0043085c
'0070e52c    56                      push esi
'0070e52d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070e52e    ff1580104000            call dword ptr [00401080]
'0070e534    b904000280              mov ecx, 80020004
'0070e539    894d98                  mov dword ptr [ebp-68], ecx
'0070e53c    b80a000000              mov eax, 0000000a
'0070e541    894590                  mov dword ptr [ebp-70], eax
'0070e544    894da8                  mov dword ptr [ebp-58], ecx
'0070e547    8945a0                  mov dword ptr [ebp-60], eax
'0070e54a    c78528ffffff24b07200    mov dword ptr [ebp+ffffff28], 0072b024
'0070e554    c78520ffffff08400000    mov dword ptr [ebp+ffffff20], 00004008
'0070e55e    6814084300              push 00430814
'0070e563    8b55e0                  mov edx, dword ptr [ebp-20]
'0070e566    52                      push edx

' *** Reference to "__vbaStrCat"
'0070e567    8b3570104000            mov esi, dword ptr [00401070]
'0070e56d    ffd6                    call esi
var_11 = ("L'erreur suivante s'est produite : ") & (var_3)
'0070e56f    8bd0                    mov edx, eax
'0070e571    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'0070e574    8b3dd0124000            mov edi, dword ptr [004012d0]
'0070e57a    ffd7                    call edi
'0070e57c    50                      push eax
'0070e57d    6870084300              push 00430870
'0070e582    ffd6                    call esi
var_76 = (var_11) & (vbCrLf)
'0070e584    8bd0                    mov edx, eax
'0070e586    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0070e589    ffd7                    call edi
'0070e58b    50                      push eax
'0070e58c    6870084300              push 00430870
'0070e591    ffd6                    call esi
var_12 = (var_76) & (vbCrLf)
'0070e593    8bd0                    mov edx, eax
'0070e595    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0070e598    ffd7                    call edi
'0070e59a    50                      push eax
'0070e59b    6880084300              push 00430880
'0070e5a0    ffd6                    call esi
var_13 = (var_12) & (" numéro d'erreur (")
'0070e5a2    8bd0                    mov edx, eax
'0070e5a4    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0070e5a7    ffd7                    call edi
'0070e5a9    50                      push eax
'0070e5aa    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'0070e5b0    50                      push eax

' *** Reference to "__vbaStrI4"
'0070e5b1    ff1520104000            call dword ptr [00401020]
'0070e5b7    8bd0                    mov edx, eax
'0070e5b9    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070e5bc    ffd7                    call edi
'0070e5be    50                      push eax
'0070e5bf    ffd6                    call esi
var_127 = (var_13) & (CStr(var_10))
'0070e5c1    8bd0                    mov edx, eax
'0070e5c3    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0070e5c6    ffd7                    call edi
'0070e5c8    50                      push eax
'0070e5c9    68ac084300              push 004308ac
'0070e5ce    ffd6                    call esi
var_15 = (var_127) & (" )")
'0070e5d0    8945b8                  mov dword ptr [ebp-48], eax
'0070e5d3    bf08000000              mov edi, 00000008
'0070e5d8    897db0                  mov dword ptr [ebp-50], edi
'0070e5db    8d4d90                  lea ecx, dword ptr [ebp-70]
'0070e5de    51                      push ecx
'0070e5df    8d55a0                  lea edx, dword ptr [ebp-60]
'0070e5e2    52                      push edx
'0070e5e3    8d8520ffffff            lea eax, dword ptr [ebp+ffffff20]
'0070e5e9    50                      push eax
'0070e5ea    6a10                    push 10
'0070e5ec    8d4db0                  lea ecx, dword ptr [ebp-50]
'0070e5ef    51                      push ecx

' *** Reference to "rtcMsgBox"
'0070e5f0    ff15b8104000            call dword ptr [004010b8]
var_128 = MsgBox(var_15, 16, vbNullString)
'0070e5f6    8d55c8                  lea edx, dword ptr [ebp-38]
'0070e5f9    52                      push edx
'0070e5fa    8d45cc                  lea eax, dword ptr [ebp-34]
'0070e5fd    50                      push eax
'0070e5fe    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0070e601    51                      push ecx
'0070e602    8d55d4                  lea edx, dword ptr [ebp-2c]
'0070e605    52                      push edx
'0070e606    8d45d8                  lea eax, dword ptr [ebp-28]
'0070e609    50                      push eax
'0070e60a    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0070e60d    51                      push ecx
'0070e60e    8d55e0                  lea edx, dword ptr [ebp-20]
'0070e611    52                      push edx
'0070e612    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0070e614    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4508, -4512, -4516, -4520, -4524, -4528)
'0070e61a    8d45c0                  lea eax, dword ptr [ebp-40]
'0070e61d    50                      push eax
'0070e61e    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0070e621    51                      push ecx
'0070e622    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070e624    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'0070e62a    8d5590                  lea edx, dword ptr [ebp-70]
'0070e62d    52                      push edx
'0070e62e    8d45a0                  lea eax, dword ptr [ebp-60]
'0070e631    50                      push eax
'0070e632    8d4db0                  lea ecx, dword ptr [ebp-50]
'0070e635    51                      push ecx
'0070e636    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'0070e638    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8)
'0070e63e    83c43c                  add esp, 3c
'0070e641    8d55b0                  lea edx, dword ptr [ebp-50]
'0070e644    52                      push edx

' *** Reference to "rtcGetPresentDate"
'0070e645    ff15f4124000            call dword ptr [004012f4]
var_15 = Now()
'0070e64b    c78528ffffffb8084300    mov dword ptr [ebp+ffffff28], 004308b8
'0070e655    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'0070e65b    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'0070e661    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaVarDup"
'0070e664    ff1598124000            call dword ptr [00401298]
var_7 = ("dd/MM/yyyy hh:mm:ss")
'0070e66a    6a01                    push 01
'0070e66c    6a01                    push 01
'0070e66e    8d45a0                  lea eax, dword ptr [ebp-60]
'0070e671    50                      push eax
'0070e672    8d4db0                  lea ecx, dword ptr [ebp-50]
'0070e675    51                      push ecx
'0070e676    8d5590                  lea edx, dword ptr [ebp-70]
'0070e679    52                      push edx

' *** Reference to "rtcVarFromFormatVar"
'0070e67a    ff1574104000            call dword ptr [00401074]
'0070e680    ffd3                    call ebx
'0070e682    50                      push eax
'0070e683    8d45c4                  lea eax, dword ptr [ebp-3c]
'0070e686    50                      push eax

' *** Reference to "__vbaObjSet"
'0070e687    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Err
'0070e68d    8bf0                    mov esi, eax
'0070e68f    8b0e                    mov ecx, dword ptr [esi]
'0070e691    8d95dcfeffff            lea edx, dword ptr [ebp+fffffedc]
'0070e697    52                      push edx
'0070e698    56                      push esi
'0070e699    ff511c                  call dword ptr [ecx+1c]
var_10 = var_9.Number
'0070e69c    dbe2                    fnclex
'0070e69e    85c0                    test ax, ax
'0070e6a0    7d0f                    jge 70e6b1
'0070e6a2    6a1c                    push 1c
'0070e6a4    685c084300              push 0043085c
'0070e6a9    56                      push esi
'0070e6aa    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070e6ab    ff1580104000            call dword ptr [00401080]
'0070e6b1    ffd3                    call ebx
'0070e6b3    50                      push eax
'0070e6b4    8d45c0                  lea eax, dword ptr [ebp-40]
'0070e6b7    50                      push eax

' *** Reference to "__vbaObjSet"
'0070e6b8    ff15b4104000            call dword ptr [004010b4]
Set var_5 = Err
'0070e6be    8bf0                    mov esi, eax
'0070e6c0    8b0e                    mov ecx, dword ptr [esi]
'0070e6c2    8d55e0                  lea edx, dword ptr [ebp-20]
'0070e6c5    52                      push edx
'0070e6c6    56                      push esi
'0070e6c7    ff512c                  call dword ptr [ecx+2c]
var_3 = var_5.Description
'0070e6ca    dbe2                    fnclex
'0070e6cc    85c0                    test ax, ax
'0070e6ce    7d0f                    jge 70e6df
'0070e6d0    6a2c                    push 2c
'0070e6d2    685c084300              push 0043085c
'0070e6d7    56                      push esi
'0070e6d8    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070e6d9    ff1580104000            call dword ptr [00401080]
'0070e6df    c78518ffffffe4084300    mov dword ptr [ebp+ffffff18], 004308e4
'0070e6e9    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'0070e6ef    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'0070e6f5    898508ffffff            mov dword ptr [ebp+ffffff08], eax
'0070e6fb    c78500ffffff03000000    mov dword ptr [ebp+ffffff00], 00000003
'0070e705    c785f8feffff08094300    mov dword ptr [ebp+fffffef8], 00430908
'0070e70f    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'0070e715    8b45e0                  mov eax, dword ptr [ebp-20]
'0070e718    c745e000000000          mov dword ptr [ebp-20], 00000000
'0070e71f    898558ffffff            mov dword ptr [ebp+ffffff58], eax
'0070e725    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'0070e72b    c785e8feffffe4654500    mov dword ptr [ebp+fffffee8], 004565e4
'0070e735    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'0070e73b    8d4d90                  lea ecx, dword ptr [ebp-70]
'0070e73e    51                      push ecx
'0070e73f    8d9510ffffff            lea edx, dword ptr [ebp+ffffff10]
'0070e745    52                      push edx
'0070e746    8d4580                  lea eax, dword ptr [ebp-80]
'0070e749    50                      push eax

' *** Reference to "__vbaVarCat"
'0070e74a    8b3508124000            mov esi, dword ptr [00401208]
'0070e750    ffd6                    call esi
'0070e752    50                      push eax
'0070e753    8d8d00ffffff            lea ecx, dword ptr [ebp+ffffff00]
'0070e759    51                      push ecx
'0070e75a    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0070e760    52                      push edx
'0070e761    ffd6                    call esi
'0070e763    50                      push eax
'0070e764    8d85f0feffff            lea eax, dword ptr [ebp+fffffef0]
'0070e76a    50                      push eax
'0070e76b    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'0070e771    51                      push ecx
'0070e772    ffd6                    call esi
'0070e774    50                      push eax
'0070e775    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'0070e77b    52                      push edx
'0070e77c    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'0070e782    50                      push eax
'0070e783    ffd6                    call esi
'0070e785    50                      push eax
'0070e786    8d8de0feffff            lea ecx, dword ptr [ebp+fffffee0]
'0070e78c    51                      push ecx
'0070e78d    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'0070e793    52                      push edx
'0070e794    ffd6                    call esi
'0070e796    50                      push eax
'0070e797    33c0                    xor eax, eax
'0070e799    66a12ab07200            mov eax, dword ptr [0072b02a]
'0070e79f    50                      push eax
'0070e7a0    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'0070e7a5    ff15b8114000            call dword ptr [004011b8]
Print #0, 
'0070e7ab    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0070e7ae    51                      push ecx
'0070e7af    8d55c4                  lea edx, dword ptr [ebp-3c]
'0070e7b2    52                      push edx
'0070e7b3    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070e7b5    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'0070e7bb    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'0070e7c1    50                      push eax
'0070e7c2    8d8d40ffffff            lea ecx, dword ptr [ebp+ffffff40]
'0070e7c8    51                      push ecx
'0070e7c9    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'0070e7cf    52                      push edx
'0070e7d0    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'0070e7d6    50                      push eax
'0070e7d7    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'0070e7dd    51                      push ecx
'0070e7de    8d5580                  lea edx, dword ptr [ebp-80]
'0070e7e1    52                      push edx
'0070e7e2    8d4590                  lea eax, dword ptr [ebp-70]
'0070e7e5    50                      push eax
'0070e7e6    8d4da0                  lea ecx, dword ptr [ebp-60]
'0070e7e9    51                      push ecx
'0070e7ea    8d55b0                  lea edx, dword ptr [ebp-50]
'0070e7ed    52                      push edx
'0070e7ee    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0070e7f0    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'0070e7f6    83c440                  add esp, 40
'0070e7f9    6a00                    push 00

' *** Reference to "__vbaResume"
'0070e7fb    ff1568104000            call dword ptr [00401068]
'0070e801    e9bbfcffff              jmp 70e4c1
Resume handler_70E4C1
'0070e806    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0070e809    51                      push ecx
'0070e80a    8d55cc                  lea edx, dword ptr [ebp-34]
'0070e80d    52                      push edx
'0070e80e    8d45d0                  lea eax, dword ptr [ebp-30]
'0070e811    50                      push eax
'0070e812    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0070e815    51                      push ecx
'0070e816    8d55d8                  lea edx, dword ptr [ebp-28]
'0070e819    52                      push edx
'0070e81a    8d45dc                  lea eax, dword ptr [ebp-24]
'0070e81d    50                      push eax
'0070e81e    8d4de0                  lea ecx, dword ptr [ebp-20]
'0070e821    51                      push ecx
'0070e822    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0070e824    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_3, -4508, -4512, -4516, -4520, -4524, -4528)
'0070e82a    8d55c0                  lea edx, dword ptr [ebp-40]
'0070e82d    52                      push edx
'0070e82e    8d45c4                  lea eax, dword ptr [ebp-3c]
'0070e831    50                      push eax
'0070e832    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070e834    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'0070e83a    8d8d30ffffff            lea ecx, dword ptr [ebp+ffffff30]
'0070e840    51                      push ecx
'0070e841    8d9540ffffff            lea edx, dword ptr [ebp+ffffff40]
'0070e847    52                      push edx
'0070e848    8d8550ffffff            lea eax, dword ptr [ebp+ffffff50]
'0070e84e    50                      push eax
'0070e84f    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'0070e855    51                      push ecx
'0070e856    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0070e85c    52                      push edx
'0070e85d    8d4580                  lea eax, dword ptr [ebp-80]
'0070e860    50                      push eax
'0070e861    8d4d90                  lea ecx, dword ptr [ebp-70]
'0070e864    51                      push ecx
'0070e865    8d55a0                  lea edx, dword ptr [ebp-60]
'0070e868    52                      push edx
'0070e869    8d45b0                  lea eax, dword ptr [ebp-50]
'0070e86c    50                      push eax
'0070e86d    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0070e86f    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'0070e875    83c454                  add esp, 54
'0070e878    c3                      ret
'0070e879    c3                      ret
'0070e87a    8b4508                  mov eax, dword ptr [ebp+08]
'0070e87d    8b08                    mov ecx, dword ptr [eax]
'0070e87f    50                      push eax
'0070e880    ff5108                  call dword ptr [ecx+08]
'0070e883    8b45f4                  mov eax, dword ptr [ebp-0c]
'0070e886    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'0070e889    64890d00000000          mov dword ptr fs:[00000000], ecx
'0070e890    5f                      pop edi
'0070e891    5e                      pop esi
'0070e892    5b                      pop ebx
'0070e893    8be5                    mov esp, ebp
'0070e895    5d                      pop ebp
'0070e896    c20400                  ret 0004


End Sub


Private Sub Form_Resize()
'0070eab0    55                      push ebp
'0070eab1    8bec                    mov ebp, esp
'0070eab3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0070eab6    6866724000              push 00407266
'0070eabb    64a100000000            mov ax, word ptr fs:[00000000]
'0070eac1    50                      push eax
    'Reference to '__except_list'
'0070eac2    64892500000000          mov dword ptr fs:[00000000], esp
'0070eac9    81ec94000000            sub esp, 00000094
'0070eacf    53                      push ebx
'0070ead0    56                      push esi
'0070ead1    57                      push edi
'0070ead2    8965f4                  mov dword ptr [ebp-0c], esp
'0070ead5    c745f8a86c4000          mov dword ptr [ebp-08], 00406ca8
'0070eadc    8b7508                  mov esi, dword ptr [ebp+08]
'0070eadf    8bc6                    mov eax, esi
'0070eae1    83e001                  and eax, 01
'0070eae4    8945fc                  mov dword ptr [ebp-04], eax
'0070eae7    83e6fe                  and esi, fffffffe
'0070eaea    8b0e                    mov ecx, dword ptr [esi]
'0070eaec    56                      push esi
'0070eaed    897508                  mov dword ptr [ebp+08], esi
'0070eaf0    ff5104                  call dword ptr [ecx+04]
'0070eaf3    a14cb17200              mov ax, word ptr [0072b14c]
'0070eaf8    33db                    xor ebx, ebx
'0070eafa    3bc3                    cmp eax, ebx
'0070eafc    895ddc                  mov dword ptr [ebp-24], ebx
'0070eaff    895dd8                  mov dword ptr [ebp-28], ebx
'0070eb02    895dd4                  mov dword ptr [ebp-2c], ebx
'0070eb05    895dc4                  mov dword ptr [ebp-3c], ebx
'0070eb08    895db4                  mov dword ptr [ebp-4c], ebx
'0070eb0b    895d90                  mov dword ptr [ebp-70], ebx
'0070eb0e    895d8c                  mov dword ptr [ebp-74], ebx
'0070eb11    895d88                  mov dword ptr [ebp-78], ebx
'0070eb14    7510                    jne 70eb26

If (0 = 0) Then
'0070eb16    684cb17200              push 0072b14c
'0070eb1b    68c4d24000              push 0040d2c4

' *** Reference to "__vbaNew2"
'0070eb20    ff152c124000            call dword ptr [0040122c]
    Dim var_47 As New frmExperience
End If
'0070eb26    8b3d4cb17200            mov edi, dword ptr [0072b14c]
'0070eb2c    8b17                    mov edx, dword ptr [edi]
'0070eb2e    8d4590                  lea eax, dword ptr [ebp-70]
'0070eb31    50                      push eax
'0070eb32    57                      push edi
'0070eb33    ff9298000000            call dword ptr [edx+00000098]
'0070eb39    dbe2                    fnclex
'0070eb3b    3bc3                    cmp eax, ebx
'0070eb3d    7d16                    jge 70eb55

' *** Reference to "__vbaHresultCheckObj"
'0070eb3f    8b1d80104000            mov ebx, dword ptr [00401080]
'0070eb45    6898000000              push 00000098
'0070eb4a    68141d4300              push 00431d14
'0070eb4f    57                      push edi
'0070eb50    50                      push eax
'0070eb51    ffd3                    call ebx
'0070eb53    eb06                    jmp 70eb5b

' *** Reference to "__vbaHresultCheckObj"
'0070eb55    8b1d80104000            mov ebx, dword ptr [00401080]
'0070eb5b    66837d9002              cmp word ptr [ebp-70], 02
'0070eb60    0f845e010000            je 70ecc4

If (0 <> 2) Then
'0070eb66    a14cb17200              mov ax, word ptr [0072b14c]
'0070eb6b    85c0                    test ax, ax
'0070eb6d    7510                    jne 70eb7f
'0070eb6f    684cb17200              push 0072b14c
'0070eb74    68c4d24000              push 0040d2c4

' *** Reference to "__vbaNew2"
'0070eb79    ff152c124000            call dword ptr [0040122c]
    Set var_47 = New frmExperience
'0070eb7f    8b3d4cb17200            mov edi, dword ptr [0072b14c]
'0070eb85    8b0f                    mov ecx, dword ptr [edi]
'0070eb87    6800d03646              push 4636d000
'0070eb8c    57                      push edi
'0070eb8d    ff9184000000            call dword ptr [ecx+00000084]
'0070eb93    dbe2                    fnclex
'0070eb95    85c0                    test ax, ax
'0070eb97    7d0e                    jge 70eba7
'0070eb99    6884000000              push 00000084
'0070eb9e    68141d4300              push 00431d14
'0070eba3    57                      push edi
'0070eba4    50                      push eax
'0070eba5    ffd3                    call ebx
'0070eba7    a14cb17200              mov ax, word ptr [0072b14c]
'0070ebac    85c0                    test ax, ax
'0070ebae    7510                    jne 70ebc0
'0070ebb0    684cb17200              push 0072b14c
'0070ebb5    68c4d24000              push 0040d2c4

' *** Reference to "__vbaNew2"
'0070ebba    ff152c124000            call dword ptr [0040122c]
    Set var_47 = New frmExperience
'0070ebc0    8b3d4cb17200            mov edi, dword ptr [0072b14c]
'0070ebc6    8b17                    mov edx, dword ptr [edi]
'0070ebc8    8d458c                  lea eax, dword ptr [ebp-74]
'0070ebcb    50                      push eax
'0070ebcc    57                      push edi
'0070ebcd    ff9288000000            call dword ptr [edx+00000088]
'0070ebd3    dbe2                    fnclex
'0070ebd5    85c0                    test ax, ax
'0070ebd7    7d0e                    jge 70ebe7
'0070ebd9    6888000000              push 00000088
'0070ebde    68141d4300              push 00431d14
'0070ebe3    57                      push edi
'0070ebe4    50                      push eax
'0070ebe5    ffd3                    call ebx
'0070ebe7    a14cb17200              mov ax, word ptr [0072b14c]
'0070ebec    85c0                    test ax, ax
'0070ebee    7510                    jne 70ec00
'0070ebf0    684cb17200              push 0072b14c
'0070ebf5    68c4d24000              push 0040d2c4

' *** Reference to "__vbaNew2"
'0070ebfa    ff152c124000            call dword ptr [0040122c]
    Set var_47 = New frmExperience
'0070ec00    a110b07200              mov ax, word ptr [0072b010]
'0070ec05    85c0                    test ax, ax
'0070ec07    8b1d4cb17200            mov ebx, dword ptr [0072b14c]
'0070ec0d    7510                    jne 70ec1f
'0070ec0f    6810b07200              push 0072b010
'0070ec14    68b0e54000              push 0040e5b0

' *** Reference to "__vbaNew2"
'0070ec19    ff152c124000            call dword ptr [0040122c]
    Dim var_60 As New frmMain
'0070ec1f    8b3d10b07200            mov edi, dword ptr [0072b010]
'0070ec25    8b0f                    mov ecx, dword ptr [edi]
'0070ec27    8d5588                  lea edx, dword ptr [ebp-78]
'0070ec2a    52                      push edx
'0070ec2b    57                      push edi
'0070ec2c    ff9188000000            call dword ptr [ecx+00000088]
'0070ec32    dbe2                    fnclex
'0070ec34    85c0                    test ax, ax
'0070ec36    7d12                    jge 70ec4a
'0070ec38    6888000000              push 00000088
'0070ec3d    684cfd4200              push 0042fd4c
'0070ec42    57                      push edi
'0070ec43    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070ec44    ff1580104000            call dword ptr [00401080]
'0070ec4a    8b4588                  mov eax, dword ptr [ebp-78]
'0070ec4d    8b4d8c                  mov ecx, dword ptr [ebp-74]
'0070ec50    8945bc                  mov dword ptr [ebp-44], eax
'0070ec53    b804000000              mov eax, 00000004
'0070ec58    8945b4                  mov dword ptr [ebp-4c], eax
'0070ec5b    8945c4                  mov dword ptr [ebp-3c], eax
'0070ec5e    8d55b4                  lea edx, dword ptr [ebp-4c]
'0070ec61    52                      push edx
'0070ec62    8d45c4                  lea eax, dword ptr [ebp-3c]
'0070ec65    894dcc                  mov dword ptr [ebp-34], ecx
'0070ec68    8b3b                    mov edi, dword ptr [ebx]
'0070ec6a    50                      push eax
'0070ec6b    e8203bdeff              call 4f2790
    Call sub_4F2790()
'0070ec70    0fbfc8                  movsx ecx, ax
'0070ec73    898d64ffffff            mov dword ptr [ebp+ffffff64], ecx
'0070ec79    db8564ffffff            fild dword ptr [ebp+ffffff64]
'0070ec7f    d99d60ffffff            fstp dword ptr [ebp+ffffff60]
    'var_20 = (00)
'0070ec85    8b9560ffffff            mov edx, dword ptr [ebp+ffffff60]
'0070ec8b    52                      push edx
'0070ec8c    53                      push ebx
'0070ec8d    ff978c000000            call dword ptr [edi+0000008c]
'0070ec93    dbe2                    fnclex
'0070ec95    85c0                    test ax, ax
'0070ec97    7d12                    jge 70ecab
'0070ec99    688c000000              push 0000008c
'0070ec9e    68141d4300              push 00431d14
'0070eca3    53                      push ebx
'0070eca4    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070eca5    ff1580104000            call dword ptr [00401080]
'0070ecab    8d45b4                  lea eax, dword ptr [ebp-4c]
'0070ecae    50                      push eax
'0070ecaf    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0070ecb2    51                      push ecx
'0070ecb3    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'0070ecb5    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_9, var_62)

' *** Reference to "__vbaHresultCheckObj"
'0070ecbb    8b1d80104000            mov ebx, dword ptr [00401080]
'0070ecc1    83c40c                  add esp, 0c
    
End If
'0070ecc4    a14cb17200              mov ax, word ptr [0072b14c]
'0070ecc9    85c0                    test ax, ax
'0070eccb    7510                    jne 70ecdd
'0070eccd    684cb17200              push 0072b14c
'0070ecd2    68c4d24000              push 0040d2c4

' *** Reference to "__vbaNew2"
'0070ecd7    ff152c124000            call dword ptr [0040122c]
Set var_47 = New frmExperience
'0070ecdd    8b3d4cb17200            mov edi, dword ptr [0072b14c]
'0070ece3    8b17                    mov edx, dword ptr [edi]
'0070ece5    8d458c                  lea eax, dword ptr [ebp-74]
'0070ece8    50                      push eax
'0070ece9    57                      push edi
'0070ecea    ff9288000000            call dword ptr [edx+00000088]
'0070ecf0    dbe2                    fnclex
'0070ecf2    85c0                    test ax, ax
'0070ecf4    7d0e                    jge 70ed04
'0070ecf6    6888000000              push 00000088
'0070ecfb    68141d4300              push 00431d14
'0070ed00    57                      push edi
'0070ed01    50                      push eax
'0070ed02    ffd3                    call ebx
'0070ed04    d9458c                  fld dword ptr [ebp-74]
'0070ed07    8b55a8                  mov edx, dword ptr [ebp-58]
'0070ed0a    d825a06c4000            fsub dword ptr [00406ca0]
'0070ed10    dfe0                    fnstsw ax
'0070ed12    a80d                    test al, 0d
'0070ed14    0f8511030000            jne 70f02b
'0070ed1a    dc0d20674000            fmul qword ptr [00406720]
'0070ed20    d95dac                  fstp dword ptr [ebp-54]
'var_50 = (00)
'0070ed23    dfe0                    fnstsw ax
'0070ed25    a80d                    test al, 0d
'0070ed27    0f85fe020000            jne 70f02b
'0070ed2d    83ec10                  sub esp, 10
'0070ed30    8bcc                    mov ecx, esp
'0070ed32    b804000000              mov eax, 00000004
'0070ed37    8901                    mov dword ptr [ecx], eax
'0070ed39    8b45ac                  mov eax, dword ptr [ebp-54]
'0070ed3c    895104                  mov dword ptr [ecx+04], edx
'0070ed3f    8b55b0                  mov edx, dword ptr [ebp-50]
'0070ed42    894108                  mov dword ptr [ecx+08], eax
'0070ed45    8b06                    mov eax, dword ptr [esi]
'0070ed47    6806000180              push 80010006
'0070ed4c    56                      push esi
'0070ed4d    89510c                  mov dword ptr [ecx+0c], edx
'0070ed50    ff9020030000            call dword ptr [eax+00000320]

' *** Reference to "__vbaObjSet"
'0070ed56    8b3db4104000            mov edi, dword ptr [004010b4]
'0070ed5c    50                      push eax
'0070ed5d    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0070ed60    51                      push ecx
'0070ed61    ffd7                    call edi
Set var_45 = Nothing
'0070ed63    50                      push eax

' *** Reference to "__vbaLateIdSt"
'0070ed64    ff15ec124000            call dword ptr [004012ec]
var_45.[UNMANAGED] = (((0) - 0!) * 0.6#)
'0070ed6a    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'0070ed6d    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'0070ed73    8b16                    mov edx, dword ptr [esi]
'0070ed75    56                      push esi
'0070ed76    ff920c030000            call dword ptr [edx+0000030c]
'0070ed7c    50                      push eax
'0070ed7d    8d45d8                  lea eax, dword ptr [ebp-28]
'0070ed80    50                      push eax
'0070ed81    ffd7                    call edi
Set var_45 = Nothing
'0070ed83    89857cffffff            mov dword ptr [ebp+ffffff7c], eax
'0070ed89    a14cb17200              mov ax, word ptr [0072b14c]
'0070ed8e    85c0                    test ax, ax
'0070ed90    7510                    jne 70eda2
'0070ed92    684cb17200              push 0072b14c
'0070ed97    68c4d24000              push 0040d2c4

' *** Reference to "__vbaNew2"
'0070ed9c    ff152c124000            call dword ptr [0040122c]
Set var_47 = New frmExperience
'0070eda2    8b1d4cb17200            mov ebx, dword ptr [0072b14c]
'0070eda8    8b0b                    mov ecx, dword ptr [ebx]
'0070edaa    8d558c                  lea edx, dword ptr [ebp-74]
'0070edad    52                      push edx
'0070edae    53                      push ebx
'0070edaf    ff9188000000            call dword ptr [ecx+00000088]
'0070edb5    dbe2                    fnclex
'0070edb7    85c0                    test ax, ax
'0070edb9    7d12                    jge 70edcd
'0070edbb    6888000000              push 00000088
'0070edc0    68141d4300              push 00431d14
'0070edc5    53                      push ebx
'0070edc6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070edc7    ff1580104000            call dword ptr [00401080]
'0070edcd    d9458c                  fld dword ptr [ebp-74]
'0070edd0    8b9d7cffffff            mov ebx, dword ptr [ebp+ffffff7c]
'0070edd6    d825a06c4000            fsub dword ptr [00406ca0]
'0070eddc    8b0b                    mov ecx, dword ptr [ebx]
'0070edde    51                      push ecx
'0070eddf    dfe0                    fnstsw ax
'0070ede1    a80d                    test al, 0d
'0070ede3    0f8542020000            jne 70f02b
'0070ede9    dc0d986c4000            fmul qword ptr [00406c98]
'0070edef    dfe0                    fnstsw ax
'0070edf1    a80d                    test al, 0d
'0070edf3    0f8532020000            jne 70f02b
'0070edf9    d99d5cffffff            fstp dword ptr [ebp+ffffff5c]
'var_88 = (00)
'0070edff    d9855cffffff            fld dword ptr [ebp+ffffff5c]
'0070ee05    d91c24                  fstp dword ptr [esp]
'var_22 = (00)
'0070ee08    53                      push ebx
'0070ee09    ff9184000000            call dword ptr [ecx+00000084]
'0070ee0f    dbe2                    fnclex
'0070ee11    85c0                    test ax, ax
'0070ee13    7d12                    jge 70ee27
'0070ee15    6884000000              push 00000084
'0070ee1a    68d00d4300              push 00430dd0
'0070ee1f    53                      push ebx
'0070ee20    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070ee21    ff1580104000            call dword ptr [00401080]
'0070ee27    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'0070ee2a    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'0070ee30    8b16                    mov edx, dword ptr [esi]
'0070ee32    56                      push esi
'0070ee33    ff920c030000            call dword ptr [edx+0000030c]
'0070ee39    50                      push eax
'0070ee3a    8d45d4                  lea eax, dword ptr [ebp-2c]
'0070ee3d    50                      push eax
'0070ee3e    ffd7                    call edi
Set var_44 = 
'0070ee40    8b0e                    mov ecx, dword ptr [esi]
'0070ee42    8b18                    mov ebx, dword ptr [eax]
'0070ee44    6a00                    push 00
'0070ee46    6806000180              push 80010006
'0070ee4b    56                      push esi
'0070ee4c    894584                  mov dword ptr [ebp-7c], eax
'0070ee4f    ff9120030000            call dword ptr [ecx+00000320]
'0070ee55    50                      push eax
'0070ee56    8d55d8                  lea edx, dword ptr [ebp-28]
'0070ee59    52                      push edx
'0070ee5a    ffd7                    call edi
Set var_45 = var_44
'0070ee5c    50                      push eax
'0070ee5d    8d45c4                  lea eax, dword ptr [ebp-3c]
'0070ee60    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'0070ee61    ff157c114000            call dword ptr [0040117c]
var_9 = var_45.UNK_frmExperience_6
'0070ee67    83c410                  add esp, 10
'0070ee6a    50                      push eax

' *** Reference to "__vbaR4Var"
'0070ee6b    ff1578114000            call dword ptr [00401178]
'0070ee71    d80504674000            fadd dword ptr [00406704]
'0070ee77    51                      push ecx
'0070ee78    8bcb                    mov ecx, ebx
'0070ee7a    8b5d84                  mov ebx, dword ptr [ebp-7c]
'0070ee7d    dfe0                    fnstsw ax
'0070ee7f    a80d                    test al, 0d
'0070ee81    0f85a4010000            jne 70f02b
'0070ee87    d91c24                  fstp dword ptr [esp]
'var_89 = (00)
'0070ee8a    53                      push ebx
'0070ee8b    ff5174                  call dword ptr [ecx+74]
'0070ee8e    dbe2                    fnclex
'0070ee90    85c0                    test ax, ax
'0070ee92    7d0f                    jge 70eea3
'0070ee94    6a74                    push 74
'0070ee96    68d00d4300              push 00430dd0
'0070ee9b    53                      push ebx
'0070ee9c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070ee9d    ff1580104000            call dword ptr [00401080]
'0070eea3    8d55d4                  lea edx, dword ptr [ebp-2c]
'0070eea6    52                      push edx
'0070eea7    8d45d8                  lea eax, dword ptr [ebp-28]
'0070eeaa    50                      push eax
'0070eeab    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070eead    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_45, var_44)
'0070eeb3    83c40c                  add esp, 0c
'0070eeb6    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeVar"
'0070eeb9    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_9)
'0070eebf    8b0e                    mov ecx, dword ptr [esi]
'0070eec1    56                      push esi
'0070eec2    ff9110030000            call dword ptr [ecx+00000310]
'0070eec8    50                      push eax
'0070eec9    8d55d8                  lea edx, dword ptr [ebp-28]
'0070eecc    52                      push edx
'0070eecd    ffd7                    call edi
Set var_45 = 
'0070eecf    89857cffffff            mov dword ptr [ebp+ffffff7c], eax
'0070eed5    a14cb17200              mov ax, word ptr [0072b14c]
'0070eeda    85c0                    test ax, ax
'0070eedc    7510                    jne 70eeee
'0070eede    684cb17200              push 0072b14c
'0070eee3    68c4d24000              push 0040d2c4

' *** Reference to "__vbaNew2"
'0070eee8    ff152c124000            call dword ptr [0040122c]
Set var_47 = New frmExperience
'0070eeee    8b1d4cb17200            mov ebx, dword ptr [0072b14c]
'0070eef4    8b03                    mov eax, dword ptr [ebx]
'0070eef6    8d4d8c                  lea ecx, dword ptr [ebp-74]
'0070eef9    51                      push ecx
'0070eefa    53                      push ebx
'0070eefb    ff9088000000            call dword ptr [eax+00000088]
'0070ef01    dbe2                    fnclex
'0070ef03    85c0                    test ax, ax
'0070ef05    7d12                    jge 70ef19
'0070ef07    6888000000              push 00000088
'0070ef0c    68141d4300              push 00431d14
'0070ef11    53                      push ebx
'0070ef12    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070ef13    ff1580104000            call dword ptr [00401080]
'0070ef19    d9458c                  fld dword ptr [ebp-74]
'0070ef1c    8b9d7cffffff            mov ebx, dword ptr [ebp+ffffff7c]
'0070ef22    d8259c624000            fsub dword ptr [0040629c]
'0070ef28    8b13                    mov edx, dword ptr [ebx]
'0070ef2a    51                      push ecx
'0070ef2b    dfe0                    fnstsw ax
'0070ef2d    a80d                    test al, 0d
'0070ef2f    0f85f6000000            jne 70f02b
'0070ef35    d91c24                  fstp dword ptr [esp]
'var_21 = (00)
'0070ef38    53                      push ebx
'0070ef39    ff5274                  call dword ptr [edx+74]
'0070ef3c    dbe2                    fnclex
'0070ef3e    85c0                    test ax, ax
'0070ef40    7d0f                    jge 70ef51
'0070ef42    6a74                    push 74
'0070ef44    6838eb4300              push 0043eb38
'0070ef49    53                      push ebx
'0070ef4a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070ef4b    ff1580104000            call dword ptr [00401080]
'0070ef51    8d4dd8                  lea ecx, dword ptr [ebp-28]

' *** Reference to "__vbaFreeObj"
'0070ef54    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_45)
'0070ef5a    8b06                    mov eax, dword ptr [esi]
'0070ef5c    56                      push esi
'0070ef5d    ff9008030000            call dword ptr [eax+00000308]
'0070ef63    50                      push eax
'0070ef64    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0070ef67    51                      push ecx
'0070ef68    ffd7                    call edi
Set var_44 = Nothing
'0070ef6a    8b16                    mov edx, dword ptr [esi]
'0070ef6c    56                      push esi
'0070ef6d    8bd8                    mov ebx, eax
'0070ef6f    ff9210030000            call dword ptr [edx+00000310]
'0070ef75    50                      push eax
'0070ef76    8d45d8                  lea eax, dword ptr [ebp-28]
'0070ef79    50                      push eax
'0070ef7a    ffd7                    call edi
Set var_45 = Nothing
'0070ef7c    8d558c                  lea edx, dword ptr [ebp-74]
'0070ef7f    8bf0                    mov esi, eax
'0070ef81    8b0e                    mov ecx, dword ptr [esi]
'0070ef83    52                      push edx
'0070ef84    56                      push esi
'0070ef85    ff5170                  call dword ptr [ecx+70]
'0070ef88    dbe2                    fnclex
'0070ef8a    85c0                    test ax, ax
'0070ef8c    7d0f                    jge 70ef9d
'0070ef8e    6a70                    push 70
'0070ef90    6838eb4300              push 0043eb38
'0070ef95    56                      push esi
'0070ef96    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070ef97    ff1580104000            call dword ptr [00401080]
'0070ef9d    8b4d8c                  mov ecx, dword ptr [ebp-74]
'0070efa0    8b03                    mov eax, dword ptr [ebx]
'0070efa2    51                      push ecx
'0070efa3    53                      push ebx
'0070efa4    ff5074                  call dword ptr [eax+74]
'0070efa7    dbe2                    fnclex
'0070efa9    85c0                    test ax, ax
'0070efab    7d0f                    jge 70efbc
'0070efad    6a74                    push 74
'0070efaf    6838eb4300              push 0043eb38
'0070efb4    53                      push ebx
'0070efb5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070efb6    ff1580104000            call dword ptr [00401080]
'0070efbc    8d55d4                  lea edx, dword ptr [ebp-2c]
'0070efbf    52                      push edx
'0070efc0    8d45d8                  lea eax, dword ptr [ebp-28]
'0070efc3    50                      push eax
'0070efc4    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070efc6    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_45, var_44)
'0070efcc    83c40c                  add esp, 0c
'0070efcf    c745fc00000000          mov dword ptr [ebp-04], 00000000
'0070efd6    9b                      fwait
'0070efd7    680cf07000              push 0070f00c
'0070efdc    eb24                    jmp 70f002
'0070efde    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0070efe1    51                      push ecx
'0070efe2    8d55d8                  lea edx, dword ptr [ebp-28]
'0070efe5    52                      push edx
'0070efe6    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070efe8    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_45, var_44)
'0070efee    8d45b4                  lea eax, dword ptr [ebp-4c]
'0070eff1    50                      push eax
'0070eff2    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0070eff5    51                      push ecx
'0070eff6    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'0070eff8    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_9, var_62)
'0070effe    83c418                  add esp, 18
'0070f001    c3                      ret
'0070f002    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeVar"
'0070f005    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_39)
'0070f00b    c3                      ret
'0070f00c    8b4508                  mov eax, dword ptr [ebp+08]
'0070f00f    8b10                    mov edx, dword ptr [eax]
'0070f011    50                      push eax
'0070f012    ff5208                  call dword ptr [edx+08]
'0070f015    8b45fc                  mov eax, dword ptr [ebp-04]
'0070f018    8b4dec                  mov ecx, dword ptr [ebp-14]
'0070f01b    5f                      pop edi
'0070f01c    5e                      pop esi
    'Reference to '__except_list'
'0070f01d    64890d00000000          mov dword ptr fs:[00000000], ecx
'0070f024    5b                      pop ebx
'0070f025    8be5                    mov esp, ebp
'0070f027    5d                      pop ebp
'0070f028    c20400                  ret 0004


End Sub


Private Sub Form_MouseDown(Button as Integer, Shift as Integer, X as Single, Y as Single)
'0070e8a0    55                      push ebp
'0070e8a1    8bec                    mov ebp, esp
'0070e8a3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0070e8a6    6866724000              push 00407266
'0070e8ab    64a100000000            mov ax, word ptr fs:[00000000]
'0070e8b1    50                      push eax
    'Reference to '__except_list'
'0070e8b2    64892500000000          mov dword ptr fs:[00000000], esp
'0070e8b9    81ec9c000000            sub esp, 0000009c
'0070e8bf    53                      push ebx
'0070e8c0    56                      push esi
'0070e8c1    57                      push edi
'0070e8c2    8965f4                  mov dword ptr [ebp-0c], esp
'0070e8c5    c745f8886c4000          mov dword ptr [ebp-08], 00406c88
'0070e8cc    8b4508                  mov eax, dword ptr [ebp+08]
'0070e8cf    8bc8                    mov ecx, eax
'0070e8d1    83e101                  and ecx, 01
'0070e8d4    894dfc                  mov dword ptr [ebp-04], ecx
'0070e8d7    83e0fe                  and eax, fffffffe
'0070e8da    8b10                    mov edx, dword ptr [eax]
'0070e8dc    50                      push eax
'0070e8dd    894508                  mov dword ptr [ebp+08], eax
'0070e8e0    ff5204                  call dword ptr [edx+04]
'0070e8e3    33db                    xor ebx, ebx
'0070e8e5    b80a000000              mov eax, 0000000a
'0070e8ea    b904000280              mov ecx, 80020004
'0070e8ef    895da8                  mov dword ptr [ebp-58], ebx
'0070e8f2    895d98                  mov dword ptr [ebp-68], ebx
'0070e8f5    895d88                  mov dword ptr [ebp-78], ebx
'0070e8f8    894d90                  mov dword ptr [ebp-70], ecx
'0070e8fb    894588                  mov dword ptr [ebp-78], eax
'0070e8fe    894da0                  mov dword ptr [ebp-60], ecx
'0070e901    894598                  mov dword ptr [ebp-68], eax
'0070e904    894db0                  mov dword ptr [ebp-50], ecx
'0070e907    8945a8                  mov dword ptr [ebp-58], eax
'0070e90a    8b4510                  mov eax, dword ptr [ebp+10]
'0070e90d    33c9                    xor ecx, ecx
var_num3 = Empty
'0070e90f    668b08                  mov cx, word ptr [eax]
'0070e912    6850664500              push 00456650
'0070e917    895de8                  mov dword ptr [ebp-18], ebx
'0070e91a    895de4                  mov dword ptr [ebp-1c], ebx
'0070e91d    895de0                  mov dword ptr [ebp-20], ebx
'0070e920    895ddc                  mov dword ptr [ebp-24], ebx
'0070e923    895dd8                  mov dword ptr [ebp-28], ebx
'0070e926    51                      push ecx
'0070e927    895dd4                  mov dword ptr [ebp-2c], ebx
'0070e92a    895dd0                  mov dword ptr [ebp-30], ebx
'0070e92d    895dcc                  mov dword ptr [ebp-34], ebx
'0070e930    895dc8                  mov dword ptr [ebp-38], ebx
'0070e933    895db8                  mov dword ptr [ebp-48], ebx

' *** Reference to "__vbaStrI2"
'0070e936    ff1510104000            call dword ptr [00401010]

' *** Reference to "__vbaStrMove"
'0070e93c    8b35d0124000            mov esi, dword ptr [004012d0]
'0070e942    8bd0                    mov edx, eax
'0070e944    8d4de8                  lea ecx, dword ptr [ebp-18]
'0070e947    ffd6                    call esi

' *** Reference to "__vbaStrCat"
'0070e949    8b3d70104000            mov edi, dword ptr [00401070]
'0070e94f    50                      push eax
'0070e950    ffd7                    call edi
var_84 = ("shift ") & (CStr(arg_1))
'0070e952    8bd0                    mov edx, eax
'0070e954    8d4de4                  lea ecx, dword ptr [ebp-1c]
'0070e957    ffd6                    call esi
'0070e959    50                      push eax
'0070e95a    6844ed4300              push 0043ed44
'0070e95f    ffd7                    call edi
var_63 = (var_84) & (vbTab)
'0070e961    8bd0                    mov edx, eax
'0070e963    8d4de0                  lea ecx, dword ptr [ebp-20]
'0070e966    ffd6                    call esi
'0070e968    50                      push eax
'0070e969    686c334500              push 0045336c
'0070e96e    ffd7                    call edi
var_36 = (var_63) & ("x ")
'0070e970    8bd0                    mov edx, eax
'0070e972    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0070e975    ffd6                    call esi
'0070e977    8b5514                  mov edx, dword ptr [ebp+14]
'0070e97a    50                      push eax
'0070e97b    8b02                    mov eax, dword ptr [edx]
'0070e97d    50                      push eax

' *** Reference to "__vbaStrR4"
'0070e97e    ff156c114000            call dword ptr [0040116c]
'0070e984    8bd0                    mov edx, eax
'0070e986    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0070e989    ffd6                    call esi
'0070e98b    50                      push eax
'0070e98c    ffd7                    call edi
var_12 = (var_36) & (CStr(arg_2))
'0070e98e    8bd0                    mov edx, eax
'0070e990    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0070e993    ffd6                    call esi
'0070e995    50                      push eax
'0070e996    6844ed4300              push 0043ed44
'0070e99b    ffd7                    call edi
var_13 = (var_12) & (vbTab)
'0070e99d    8bd0                    mov edx, eax
'0070e99f    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0070e9a2    ffd6                    call esi
'0070e9a4    50                      push eax
'0070e9a5    6864664500              push 00456664
'0070e9aa    ffd7                    call edi
var_14 = (var_13) & ("y ")
'0070e9ac    8bd0                    mov edx, eax
'0070e9ae    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070e9b1    ffd6                    call esi
'0070e9b3    8b4d18                  mov ecx, dword ptr [ebp+18]
'0070e9b6    8b11                    mov edx, dword ptr [ecx]
'0070e9b8    50                      push eax
'0070e9b9    52                      push edx

' *** Reference to "__vbaStrR4"
'0070e9ba    ff156c114000            call dword ptr [0040116c]
'0070e9c0    8bd0                    mov edx, eax
'0070e9c2    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0070e9c5    ffd6                    call esi
'0070e9c7    50                      push eax
'0070e9c8    ffd7                    call edi
var_15 = (var_14) & (CStr(arg_3))
'0070e9ca    8945c0                  mov dword ptr [ebp-40], eax
'0070e9cd    8d4588                  lea eax, dword ptr [ebp-78]
'0070e9d0    50                      push eax
'0070e9d1    8d4d98                  lea ecx, dword ptr [ebp-68]
'0070e9d4    51                      push ecx
'0070e9d5    8d55a8                  lea edx, dword ptr [ebp-58]
'0070e9d8    52                      push edx
'0070e9d9    53                      push ebx
'0070e9da    8d45b8                  lea eax, dword ptr [ebp-48]
'0070e9dd    50                      push eax
'0070e9de    c745b808000000          mov dword ptr [ebp-48], 00000008

' *** Reference to "rtcMsgBox"
'0070e9e5    ff15b8104000            call dword ptr [004010b8]
var_128 = MsgBox(var_15, 0)
'0070e9eb    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0070e9ee    51                      push ecx
'0070e9ef    8d55cc                  lea edx, dword ptr [ebp-34]
'0070e9f2    52                      push edx
'0070e9f3    8d45d0                  lea eax, dword ptr [ebp-30]
'0070e9f6    50                      push eax
'0070e9f7    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0070e9fa    51                      push ecx
'0070e9fb    8d55d8                  lea edx, dword ptr [ebp-28]
'0070e9fe    52                      push edx
'0070e9ff    8d45dc                  lea eax, dword ptr [ebp-24]
'0070ea02    50                      push eax
'0070ea03    8d4de0                  lea ecx, dword ptr [ebp-20]
'0070ea06    51                      push ecx
'0070ea07    8d55e4                  lea edx, dword ptr [ebp-1c]
'0070ea0a    52                      push edx
'0070ea0b    8d45e8                  lea eax, dword ptr [ebp-18]
'0070ea0e    50                      push eax
'0070ea0f    6a09                    push 09

' *** Reference to "__vbaFreeStrList"
'0070ea11    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4496, -4500, -4504, -4508, -4512, -4516, -4520, -4524, -4528)
'0070ea17    8d4d88                  lea ecx, dword ptr [ebp-78]
'0070ea1a    51                      push ecx
'0070ea1b    8d5598                  lea edx, dword ptr [ebp-68]
'0070ea1e    52                      push edx
'0070ea1f    8d45a8                  lea eax, dword ptr [ebp-58]
'0070ea22    50                      push eax
'0070ea23    8d4db8                  lea ecx, dword ptr [ebp-48]
'0070ea26    51                      push ecx
'0070ea27    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0070ea29    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_61, var_86, var_130, var_131)
'0070ea2f    83c43c                  add esp, 3c
'0070ea32    895dfc                  mov dword ptr [ebp-04], ebx
'0070ea35    9b                      fwait
'0070ea36    6886ea7000              push 0070ea86
'0070ea3b    eb48                    jmp 70ea85
'0070ea3d    8d55c8                  lea edx, dword ptr [ebp-38]
'0070ea40    52                      push edx
'0070ea41    8d45cc                  lea eax, dword ptr [ebp-34]
'0070ea44    50                      push eax
'0070ea45    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0070ea48    51                      push ecx
'0070ea49    8d55d4                  lea edx, dword ptr [ebp-2c]
'0070ea4c    52                      push edx
'0070ea4d    8d45d8                  lea eax, dword ptr [ebp-28]
'0070ea50    50                      push eax
'0070ea51    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0070ea54    51                      push ecx
'0070ea55    8d55e0                  lea edx, dword ptr [ebp-20]
'0070ea58    52                      push edx
'0070ea59    8d45e4                  lea eax, dword ptr [ebp-1c]
'0070ea5c    50                      push eax
'0070ea5d    8d4de8                  lea ecx, dword ptr [ebp-18]
'0070ea60    51                      push ecx
'0070ea61    6a09                    push 09

' *** Reference to "__vbaFreeStrList"
'0070ea63    ff155c124000            call dword ptr [0040125c]
'#Cleanup( -4496, -4500, -4504, -4508, -4512, -4516, -4520, -4524, -4528)
'0070ea69    8d5588                  lea edx, dword ptr [ebp-78]
'0070ea6c    52                      push edx
'0070ea6d    8d4598                  lea eax, dword ptr [ebp-68]
'0070ea70    50                      push eax
'0070ea71    8d4da8                  lea ecx, dword ptr [ebp-58]
'0070ea74    51                      push ecx
'0070ea75    8d55b8                  lea edx, dword ptr [ebp-48]
'0070ea78    52                      push edx
'0070ea79    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0070ea7b    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_61, var_86, var_130, var_131)
'0070ea81    83c43c                  add esp, 3c
'0070ea84    c3                      ret
'0070ea85    c3                      ret
'0070ea86    8b4508                  mov eax, dword ptr [ebp+08]
'0070ea89    8b08                    mov ecx, dword ptr [eax]
'0070ea8b    50                      push eax
'0070ea8c    ff5108                  call dword ptr [ecx+08]
'0070ea8f    8b45fc                  mov eax, dword ptr [ebp-04]
'0070ea92    8b4dec                  mov ecx, dword ptr [ebp-14]
'0070ea95    5f                      pop edi
'0070ea96    5e                      pop esi
    'Reference to '__except_list'
'0070ea97    64890d00000000          mov dword ptr fs:[00000000], ecx
'0070ea9e    5b                      pop ebx
'0070ea9f    8be5                    mov esp, ebp
'0070eaa1    5d                      pop ebp
'0070eaa2    c21400                  ret 0014


End Sub


'Event for btnAdd
Private Sub btnAdd_Click()
'0070bb50    55                      push ebp
'0070bb51    8bec                    mov ebp, esp
'0070bb53    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0070bb56    6866724000              push 00407266
'0070bb5b    64a100000000            mov ax, word ptr fs:[00000000]
'0070bb61    50                      push eax
    'Reference to '__except_list'
'0070bb62    64892500000000          mov dword ptr fs:[00000000], esp
'0070bb69    81ec2c020000            sub esp, 0000022c
'0070bb6f    53                      push ebx
'0070bb70    56                      push esi
'0070bb71    57                      push edi
'0070bb72    8965f4                  mov dword ptr [ebp-0c], esp
'0070bb75    c745f8306c4000          mov dword ptr [ebp-08], 00406c30
'0070bb7c    8b7508                  mov esi, dword ptr [ebp+08]
'0070bb7f    8bc6                    mov eax, esi
'0070bb81    83e001                  and eax, 01
'0070bb84    8945fc                  mov dword ptr [ebp-04], eax
'0070bb87    83e6fe                  and esi, fffffffe
'0070bb8a    8b0e                    mov ecx, dword ptr [esi]
'0070bb8c    56                      push esi
'0070bb8d    897508                  mov dword ptr [ebp+08], esi
'0070bb90    ff5104                  call dword ptr [ecx+04]
'0070bb93    8b16                    mov edx, dword ptr [esi]
'0070bb95    33ff                    xor edi, edi
'0070bb97    56                      push esi
'0070bb98    897dd4                  mov dword ptr [ebp-2c], edi
'0070bb9b    897dd0                  mov dword ptr [ebp-30], edi
'0070bb9e    897dcc                  mov dword ptr [ebp-34], edi
'0070bba1    897dc8                  mov dword ptr [ebp-38], edi
'0070bba4    897dc4                  mov dword ptr [ebp-3c], edi
'0070bba7    897dc0                  mov dword ptr [ebp-40], edi
'0070bbaa    897dbc                  mov dword ptr [ebp-44], edi
'0070bbad    897db8                  mov dword ptr [ebp-48], edi
'0070bbb0    897db4                  mov dword ptr [ebp-4c], edi
'0070bbb3    897db0                  mov dword ptr [ebp-50], edi
'0070bbb6    897dac                  mov dword ptr [ebp-54], edi
'0070bbb9    897da8                  mov dword ptr [ebp-58], edi
'0070bbbc    897da4                  mov dword ptr [ebp-5c], edi
'0070bbbf    897d94                  mov dword ptr [ebp-6c], edi
'0070bbc2    897d84                  mov dword ptr [ebp-7c], edi
'0070bbc5    89bd74ffffff            mov dword ptr [ebp+ffffff74], edi
'0070bbcb    89bd64ffffff            mov dword ptr [ebp+ffffff64], edi
'0070bbd1    89bd54ffffff            mov dword ptr [ebp+ffffff54], edi
'0070bbd7    89bd44ffffff            mov dword ptr [ebp+ffffff44], edi
'0070bbdd    89bd34ffffff            mov dword ptr [ebp+ffffff34], edi
'0070bbe3    89bd24ffffff            mov dword ptr [ebp+ffffff24], edi
'0070bbe9    89bd14ffffff            mov dword ptr [ebp+ffffff14], edi
'0070bbef    89bdf4feffff            mov dword ptr [ebp+fffffef4], edi
'0070bbf5    89bdd4feffff            mov dword ptr [ebp+fffffed4], edi
'0070bbfb    89bdb4feffff            mov dword ptr [ebp+fffffeb4], edi
'0070bc01    89bda4feffff            mov dword ptr [ebp+fffffea4], edi
'0070bc07    89bd94feffff            mov dword ptr [ebp+fffffe94], edi
'0070bc0d    89bd84feffff            mov dword ptr [ebp+fffffe84], edi
'0070bc13    89bd74feffff            mov dword ptr [ebp+fffffe74], edi
'0070bc19    89bd64feffff            mov dword ptr [ebp+fffffe64], edi
'0070bc1f    89bd44feffff            mov dword ptr [ebp+fffffe44], edi
'0070bc25    ff92fc020000            call dword ptr [edx+000002fc]

' *** Reference to "__vbaObjSet"
'0070bc2b    8b1db4104000            mov ebx, dword ptr [004010b4]
'0070bc31    50                      push eax
'0070bc32    8d45b8                  lea eax, dword ptr [ebp-48]
'0070bc35    50                      push eax
'0070bc36    ffd3                    call ebx
Set var_61 = Me
'0070bc38    8b08                    mov ecx, dword ptr [eax]
'0070bc3a    8d55d0                  lea edx, dword ptr [ebp-30]
'0070bc3d    52                      push edx
'0070bc3e    50                      push eax
'0070bc3f    898514feffff            mov dword ptr [ebp+fffffe14], eax
'0070bc45    ff91a0000000            call dword ptr [ecx+000000a0]
'0070bc4b    dbe2                    fnclex
'0070bc4d    3bc7                    cmp eax, edi
'0070bc4f    7d18                    jge 70bc69

If (var_61 < 0) Then
'0070bc51    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'0070bc57    68a0000000              push 000000a0
'0070bc5c    68d00d4300              push 00430dd0
'0070bc61    51                      push ecx
'0070bc62    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070bc63    ff1580104000            call dword ptr [00401080]
    
End If
'0070bc69    8b55d0                  mov edx, dword ptr [ebp-30]
'0070bc6c    52                      push edx

' *** Reference to "rtcR8ValFromBstr"
'0070bc6d    ff1510134000            call dword ptr [00401310]

' *** Reference to "__vbaFpR8"
'0070bc73    ff15f8104000            call dword ptr [004010f8]
'0070bc79    dc1d68154000            fcomp qword ptr [00401568]
'0070bc7f    c785c4fdffff01000000    mov dword ptr [ebp+fffffdc4], 00000001
'0070bc89    dfe0                    fnstsw ax
'0070bc8b    f6c401                  test ah, 01
'0070bc8e    7506                    jne 70bc96

If ((0# <= Val(vbNullString))) Then
'0070bc90    89bdc4fdffff            mov dword ptr [ebp+fffffdc4], edi
    
End If
'0070bc96    8b06                    mov eax, dword ptr [esi]
'0070bc98    56                      push esi
'0070bc99    ff90fc020000            call dword ptr [eax+000002fc]
'0070bc9f    50                      push eax
'0070bca0    8d4db4                  lea ecx, dword ptr [ebp-4c]
'0070bca3    51                      push ecx
'0070bca4    ffd3                    call ebx
Set var_62 = Nothing
'0070bca6    8b10                    mov edx, dword ptr [eax]
'0070bca8    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070bcab    51                      push ecx
'0070bcac    50                      push eax
'0070bcad    89850cfeffff            mov dword ptr [ebp+fffffe0c], eax
'0070bcb3    ff92a0000000            call dword ptr [edx+000000a0]
'0070bcb9    dbe2                    fnclex
'0070bcbb    3bc7                    cmp eax, edi
'0070bcbd    7d18                    jge 70bcd7

If (var_62 < 0) Then
'0070bcbf    8b950cfeffff            mov edx, dword ptr [ebp+fffffe0c]
'0070bcc5    68a0000000              push 000000a0
'0070bcca    68d00d4300              push 00430dd0
'0070bccf    52                      push edx
'0070bcd0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070bcd1    ff1580104000            call dword ptr [00401080]
    
End If
'0070bcd7    8b45cc                  mov eax, dword ptr [ebp-34]
'0070bcda    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'0070bcdb    ff1510134000            call dword ptr [00401310]

' *** Reference to "__vbaFpR8"
'0070bce1    ff15f8104000            call dword ptr [004010f8]
'0070bce7    dc1d286c4000            fcomp qword ptr [00406c28]
'0070bced    dfe0                    fnstsw ax
'0070bcef    f6c441                  test ah, 41
'0070bcf2    7507                    jne 70bcfb
'0070bcf4    b801000000              mov eax, 00000001
'0070bcf9    eb02                    jmp 70bcfd
'0070bcfb    33c0                    xor eax, eax
'0070bcfd    8b8dc4fdffff            mov ecx, dword ptr [ebp+fffffdc4]
'0070bd03    f7d9                    neg ecx
'0070bd05    f7d8                    neg eax
'0070bd07    0bc1                    or eax, ecx
var_num1 = (100# < Val(vbNullString)) Or -(0)
'0070bd09    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070bd0c    51                      push ecx
'0070bd0d    8d55d0                  lea edx, dword ptr [ebp-30]
'0070bd10    52                      push edx
'0070bd11    6a02                    push 02
'0070bd13    66898504feffff          mov word ptr [ebp+fffffe04], ax

' *** Reference to "__vbaFreeStrList"
'0070bd1a    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, 0)
'0070bd20    8d45b4                  lea eax, dword ptr [ebp-4c]
'0070bd23    50                      push eax
'0070bd24    8d4db8                  lea ecx, dword ptr [ebp-48]
'0070bd27    51                      push ecx
'0070bd28    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070bd2a    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_61, var_62)
'0070bd30    83c418                  add esp, 18
'0070bd33    6639bd04feffff          cmp word ptr [ebp+fffffe04], di
'0070bd3a    0f84b3000000            je 70bdf3

If (var_num1 <> 0) Then
'0070bd40    b904000280              mov ecx, 80020004
'0070bd45    b80a000000              mov eax, 0000000a
'0070bd4a    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'0070bd50    898d7cffffff            mov dword ptr [ebp+ffffff7c], ecx
'0070bd56    894d8c                  mov dword ptr [ebp-74], ecx
'0070bd59    8d9514ffffff            lea edx, dword ptr [ebp+ffffff14]
'0070bd5f    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0070bd62    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'0070bd68    898574ffffff            mov dword ptr [ebp+ffffff74], eax
'0070bd6e    894584                  mov dword ptr [ebp-7c], eax
'0070bd71    c7851cffffffa8634500    mov dword ptr [ebp+ffffff1c], 004563a8
'0070bd7b    c78514ffffff08000000    mov dword ptr [ebp+ffffff14], 00000008

' *** Reference to "__vbaVarDup"
'0070bd85    ff1598124000            call dword ptr [00401298]
    var_121 = ("Le % des XP d'idée doit être un réél supérieur à 0 et inférieur à 100")
'0070bd8b    8d9564ffffff            lea edx, dword ptr [ebp+ffffff64]
'0070bd91    52                      push edx
'0070bd92    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'0070bd98    50                      push eax
'0070bd99    8d4d84                  lea ecx, dword ptr [ebp-7c]
'0070bd9c    51                      push ecx
'0070bd9d    57                      push edi
'0070bd9e    8d5594                  lea edx, dword ptr [ebp-6c]
'0070bda1    52                      push edx

' *** Reference to "rtcMsgBox"
'0070bda2    ff15b8104000            call dword ptr [004010b8]
    var_63 = MsgBox(var_121, 0)
'0070bda8    8d8564ffffff            lea eax, dword ptr [ebp+ffffff64]
'0070bdae    50                      push eax
'0070bdaf    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'0070bdb5    51                      push ecx
'0070bdb6    8d5584                  lea edx, dword ptr [ebp-7c]
'0070bdb9    52                      push edx
'0070bdba    8d4594                  lea eax, dword ptr [ebp-6c]
'0070bdbd    50                      push eax
'0070bdbe    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0070bdc0    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_121, var_119, var_91, var_123)
'0070bdc6    8b0e                    mov ecx, dword ptr [esi]
'0070bdc8    83c414                  add esp, 14
'0070bdcb    56                      push esi
'0070bdcc    ff91fc020000            call dword ptr [ecx+000002fc]
'0070bdd2    50                      push eax
'0070bdd3    8d55b8                  lea edx, dword ptr [ebp-48]
'0070bdd6    52                      push edx
'0070bdd7    ffd3                    call ebx
    Set var_61 = 
'0070bdd9    8bf0                    mov esi, eax
'0070bddb    8b06                    mov eax, dword ptr [esi]
'0070bddd    56                      push esi
'0070bdde    ff9004020000            call dword ptr [eax+00000204]
'0070bde4    dbe2                    fnclex
'0070bde6    3bc7                    cmp eax, edi
'0070bde8    0f8dee0d0000            jge 70cbdc
    
    If (    var_121 < 0) Then
'0070bdee    e9d70d0000              jmp 70cbca
    
End If
'0070bdf3    8b0e                    mov ecx, dword ptr [esi]
'0070bdf5    56                      push esi
'0070bdf6    ff9104030000            call dword ptr [ecx+00000304]
'0070bdfc    50                      push eax
'0070bdfd    8d55b8                  lea edx, dword ptr [ebp-48]
'0070be00    52                      push edx
'0070be01    ffd3                    call ebx
Set var_61 = 
'0070be03    8b08                    mov ecx, dword ptr [eax]
'0070be05    8d55d0                  lea edx, dword ptr [ebp-30]
'0070be08    52                      push edx
'0070be09    50                      push eax
'0070be0a    898514feffff            mov dword ptr [ebp+fffffe14], eax
'0070be10    ff91a0000000            call dword ptr [ecx+000000a0]
'0070be16    dbe2                    fnclex
'0070be18    3bc7                    cmp eax, edi
'0070be1a    7d18                    jge 70be34

If (var_61 < 0) Then
'0070be1c    8b8d14feffff            mov ecx, dword ptr [ebp+fffffe14]
'0070be22    68a0000000              push 000000a0
'0070be27    68d00d4300              push 00430dd0
'0070be2c    51                      push ecx
'0070be2d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070be2e    ff1580104000            call dword ptr [00401080]
    
End If
'0070be34    8b55d0                  mov edx, dword ptr [ebp-30]
'0070be37    52                      push edx

' *** Reference to "rtcR8ValFromBstr"
'0070be38    ff1510134000            call dword ptr [00401310]

' *** Reference to "__vbaFpR8"
'0070be3e    ff15f8104000            call dword ptr [004010f8]
'0070be44    dc1d68154000            fcomp qword ptr [00401568]
'0070be4a    dfe0                    fnstsw ax
'0070be4c    f6c441                  test ah, 41
'0070be4f    7507                    jne 70be58
'0070be51    b801000000              mov eax, 00000001
'0070be56    eb02                    jmp 70be5a
'0070be58    33c0                    xor eax, eax
'0070be5a    f7d8                    neg eax
'0070be5c    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0070be5f    6689850cfeffff          mov word ptr [ebp+fffffe0c], ax

' *** Reference to "__vbaFreeStr"
'0070be66    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_4)
'0070be6c    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'0070be6f    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'0070be75    6639bd0cfeffff          cmp word ptr [ebp+fffffe0c], di
'0070be7c    0f849e0c0000            je 70cb20

If ((0# < Val(vbNullString))) Then
'0070be82    8b06                    mov eax, dword ptr [esi]
'0070be84    57                      push edi
'0070be85    68a7000000              push 000000a7
'0070be8a    56                      push esi
'0070be8b    ff9020030000            call dword ptr [eax+00000320]
'0070be91    50                      push eax
'0070be92    8d4db8                  lea ecx, dword ptr [ebp-48]
'0070be95    51                      push ecx
'0070be96    ffd3                    call ebx
    Set var_61 = Nothing

' *** Reference to "__vbaLateIdCallLd"
'0070be98    8b3d7c114000            mov edi, dword ptr [0040117c]
'0070be9e    50                      push eax
'0070be9f    8d5594                  lea edx, dword ptr [ebp-6c]
'0070bea2    52                      push edx
'0070bea3    ffd7                    call edi
    var_121 = var_61.UNK_-256 - 12_167
'0070bea5    83c410                  add esp, 10
'0070bea8    50                      push eax

' *** Reference to "__vbaI4Var"
'0070bea9    ff157c124000            call dword ptr [0040127c]
'0070beaf    33c9                    xor ecx, ecx
'0070beb1    85c0                    test ax, ax
'0070beb3    0f9fc1                  setg cl
'0070beb6    f7d9                    neg ecx
'0070beb8    66898d14feffff          mov word ptr [ebp+fffffe14], cx
'0070bebf    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'0070bec2    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'0070bec8    8d4d94                  lea ecx, dword ptr [ebp-6c]

' *** Reference to "__vbaFreeVar"
'0070becb    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_121)
'0070bed1    33c0                    xor eax, eax
'0070bed3    66398514feffff          cmp word ptr [ebp+fffffe14], ax
'0070beda    0f84a60b0000            je 70ca86
    
    Do While (    (0# < Val(vbNullString)))
'0070bee0    8b16                    mov edx, dword ptr [esi]
'0070bee2    50                      push eax
'0070bee3    6a07                    push 07
'0070bee5    56                      push esi
'0070bee6    8945d8                  mov dword ptr [ebp-28], eax
'0070bee9    8945dc                  mov dword ptr [ebp-24], eax
'0070beec    ff9220030000            call dword ptr [edx+00000320]
'0070bef2    50                      push eax
'0070bef3    8d45b8                  lea eax, dword ptr [ebp-48]
'0070bef6    50                      push eax
'0070bef7    ffd3                    call ebx
    Set var_61 = Nothing
'0070bef9    50                      push eax
'0070befa    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0070befd    51                      push ecx
'0070befe    ffd7                    call edi
    var_121 = var_61.UNK_-256 - 12_7
'0070bf00    83c410                  add esp, 10
'0070bf03    50                      push eax

' *** Reference to "__vbaI4Var"
'0070bf04    ff157c124000            call dword ptr [0040127c]
'0070bf0a    8bc8                    mov ecx, eax
'0070bf0c    83e901                  sub ecx, 01
    var_num3 = CLng(var_121) - 1
'0070bf0f    0f80810d0000            jo 70cc96

' *** Reference to "__vbaI2I4"
'0070bf15    ff1550114000            call dword ptr [00401150]
'0070bf1b    8d4db8                  lea ecx, dword ptr [ebp-48]
'0070bf1e    c745e801000000          mov dword ptr [ebp-18], 00000001
'0070bf25    8985ecfdffff            mov dword ptr [ebp+fffffdec], eax

' *** Reference to "__vbaFreeObj"
'0070bf2b    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'0070bf31    8d4d94                  lea ecx, dword ptr [ebp-6c]

' *** Reference to "__vbaFreeVar"
'0070bf34    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_121)
'0070bf3a    668b95ecfdffff          mov dx, word ptr [ebp+fffffdec]
'0070bf41    663955e8                cmp word ptr [ebp-18], dx
'0070bf45    6a00                    push 00
'0070bf47    6a07                    push 07
'0070bf49    0f8f0c020000            jg 70c15b
    
    Do While (    var_41 <= CInt(var_num3))
'0070bf4f    8b06                    mov eax, dword ptr [esi]
'0070bf51    56                      push esi
'0070bf52    ff9020030000            call dword ptr [eax+00000320]
'0070bf58    50                      push eax
'0070bf59    8d4db8                  lea ecx, dword ptr [ebp-48]
'0070bf5c    51                      push ecx
'0070bf5d    ffd3                    call ebx
    Set var_61 = Nothing
'0070bf5f    50                      push eax
'0070bf60    8d5594                  lea edx, dword ptr [ebp-6c]
'0070bf63    52                      push edx
'0070bf64    ffd7                    call edi
    var_121 = var_61.UNK_-256 - 12_7
'0070bf66    83c410                  add esp, 10
'0070bf69    50                      push eax

' *** Reference to "__vbaI4Var"
'0070bf6a    ff157c124000            call dword ptr [0040127c]
'0070bf70    8bc8                    mov ecx, eax
'0070bf72    83e901                  sub ecx, 01
    var_num3 = CLng(var_121) - 1
'0070bf75    0f801b0d0000            jo 70cc96

' *** Reference to "__vbaI2I4"
'0070bf7b    ff1550114000            call dword ptr [00401150]
'0070bf81    8d4db8                  lea ecx, dword ptr [ebp-48]
'0070bf84    c745e400000000          mov dword ptr [ebp-1c], 00000000
'0070bf8b    8985e4fdffff            mov dword ptr [ebp+fffffde4], eax

' *** Reference to "__vbaFreeObj"
'0070bf91    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'0070bf97    8d4d94                  lea ecx, dword ptr [ebp-6c]

' *** Reference to "__vbaFreeVar"
'0070bf9a    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_121)
'0070bfa0    8b45e4                  mov eax, dword ptr [ebp-1c]
'0070bfa3    663b85e4fdffff          cmp ax, word ptr [ebp+fffffde4]
'0070bfaa    0f8f94010000            jg 70c144
    
    Do While (    var_40 <= WORD PTR [EBP+FFFFFDE4])
'0070bfb0    83ec10                  sub esp, 10
'0070bfb3    8bd4                    mov edx, esp
'0070bfb5    b903000000              mov ecx, 00000003
'0070bfba    890a                    mov dword ptr [edx], ecx
'0070bfbc    0fbfc0                  movsx eax, ax
'0070bfbf    898d14ffffff            mov dword ptr [ebp+ffffff14], ecx
'0070bfc5    8b8d18ffffff            mov ecx, dword ptr [ebp+ffffff18]
'0070bfcb    894a04                  mov dword ptr [edx+04], ecx
'0070bfce    8b0e                    mov ecx, dword ptr [esi]
'0070bfd0    6a01                    push 01
'0070bfd2    894208                  mov dword ptr [edx+08], eax
'0070bfd5    89851cffffff            mov dword ptr [ebp+ffffff1c], eax
'0070bfdb    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'0070bfe1    68a8000000              push 000000a8
'0070bfe6    56                      push esi
'0070bfe7    89420c                  mov dword ptr [edx+0c], eax
'0070bfea    ff9120030000            call dword ptr [ecx+00000320]
'0070bff0    50                      push eax
'0070bff1    8d55b8                  lea edx, dword ptr [ebp-48]
'0070bff4    52                      push edx
'0070bff5    ffd3                    call ebx
    Set var_61 = Nothing
'0070bff7    50                      push eax
'0070bff8    8d4594                  lea eax, dword ptr [ebp-6c]
'0070bffb    50                      push eax
'0070bffc    ffd7                    call edi
    var_121 = var_61.UNK_-256 - 12_168
'0070bffe    83c420                  add esp, 20
'0070c001    50                      push eax

' *** Reference to "__vbaI4Var"
'0070c002    ff157c124000            call dword ptr [0040127c]
'0070c008    0fbf4de8                movsx ecx, word ptr [ebp-18]
'0070c00c    8bf8                    mov edi, eax
'0070c00e    2bf9                    sub edi, ecx
    var_num7 = CLng(var_121) - var_41
'0070c010    f7df                    neg edi
'0070c012    1bff                    sbb edi, edi
'0070c014    47                      inc edi
'0070c015    8d4db8                  lea ecx, dword ptr [ebp-48]
'0070c018    f7df                    neg edi

' *** Reference to "__vbaFreeObj"
'0070c01a    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'0070c020    8d4d94                  lea ecx, dword ptr [ebp-6c]

' *** Reference to "__vbaFreeVar"
'0070c023    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_121)
'0070c029    6685ff                  test edi, edi
'0070c02c    0f84f5000000            je 70c127
    
    If (    -(CBool((var_num7)))) Then
'0070c032    668b55e8                mov dx, word ptr [ebp-18]
'0070c036    83ec10                  sub esp, 10
'0070c039    8bfc                    mov edi, esp
'0070c03b    33c0                    xor eax, eax
'0070c03d    b903000000              mov ecx, 00000003
'0070c042    890f                    mov dword ptr [edi], ecx
'0070c044    898d14ffffff            mov dword ptr [ebp+ffffff14], ecx
'0070c04a    8b8d18ffffff            mov ecx, dword ptr [ebp+ffffff18]
'0070c050    894f04                  mov dword ptr [edi+04], ecx
'0070c053    894708                  mov dword ptr [edi+08], eax
'0070c056    89851cffffff            mov dword ptr [ebp+ffffff1c], eax
'0070c05c    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'0070c062    89470c                  mov dword ptr [edi+0c], eax
'0070c065    83ec10                  sub esp, 10
'0070c068    668995fcfeffff          mov word ptr [ebp+fffffefc], dx
'0070c06f    8b85fcfeffff            mov eax, dword ptr [ebp+fffffefc]
'0070c075    8bcc                    mov ecx, esp
'0070c077    ba02000000              mov edx, 00000002
'0070c07c    8911                    mov dword ptr [ecx], edx
'0070c07e    8b95f8feffff            mov edx, dword ptr [ebp+fffffef8]
'0070c084    895104                  mov dword ptr [ecx+04], edx
'0070c087    8b9500ffffff            mov edx, dword ptr [ebp+ffffff00]
'0070c08d    894108                  mov dword ptr [ecx+08], eax
'0070c090    89510c                  mov dword ptr [ecx+0c], edx
'0070c093    8b95d8feffff            mov edx, dword ptr [ebp+fffffed8]
'0070c099    83ec10                  sub esp, 10
'0070c09c    8bcc                    mov ecx, esp
'0070c09e    b802000000              mov eax, 00000002
'0070c0a3    8901                    mov dword ptr [ecx], eax
'0070c0a5    895104                  mov dword ptr [ecx+04], edx
'0070c0a8    b806000000              mov eax, 00000006
'0070c0ad    894108                  mov dword ptr [ecx+08], eax
'0070c0b0    8b85e0feffff            mov eax, dword ptr [ebp+fffffee0]
'0070c0b6    6a03                    push 03
'0070c0b8    689e000000              push 0000009e
'0070c0bd    89410c                  mov dword ptr [ecx+0c], eax
'0070c0c0    8b0e                    mov ecx, dword ptr [esi]
'0070c0c2    56                      push esi
'0070c0c3    ff9120030000            call dword ptr [ecx+00000320]
'0070c0c9    50                      push eax
'0070c0ca    8d55b8                  lea edx, dword ptr [ebp-48]
'0070c0cd    52                      push edx
'0070c0ce    ffd3                    call ebx
    Set var_61 = Nothing
'0070c0d0    50                      push eax
'0070c0d1    8d4594                  lea eax, dword ptr [ebp-6c]
'0070c0d4    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'0070c0d5    ff157c114000            call dword ptr [0040117c]
    var_121 = var_61.UNK_-256 - 12_158
'0070c0db    83c440                  add esp, 40
'0070c0de    50                      push eax
'0070c0df    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0070c0e2    51                      push ecx

' *** Reference to "__vbaStrVarVal"
'0070c0e3    ff15fc114000            call dword ptr [004011fc]
'0070c0e9    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'0070c0ea    ff1510134000            call dword ptr [00401310]
'0070c0f0    dd9d18feffff            fstp qword ptr [ebp+fffffe18]
    'var_694 = (00)
'0070c0f6    dd8518feffff            fld qword ptr [ebp+fffffe18]
'0070c0fc    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0070c0ff    dc45d8                  fadd qword ptr [ebp-28]
'0070c102    dd5dd8                  fstp qword ptr [ebp-28]
    'var_45 = (00)
'0070c105    dfe0                    fnstsw ax
'0070c107    a80d                    test al, 0d
'0070c109    0f85820b0000            jne 70cc91

' *** Reference to "__vbaFreeStr"
'0070c10f    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_4)
'0070c115    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'0070c118    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'0070c11e    8d4d94                  lea ecx, dword ptr [ebp-6c]

' *** Reference to "__vbaFreeVar"
'0070c121    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_121)
End If

' *** Reference to "__vbaLateIdCallLd"
'0070c127    8b3d7c114000            mov edi, dword ptr [0040117c]
'0070c12d    b801000000              mov eax, 00000001
'0070c132    660345e4                add ax, word ptr [ebp-1c]
var_num1 = 1 + var_40
'0070c136    0f805a0b0000            jo 70cc96
'0070c13c    8945e4                  mov dword ptr [ebp-1c], eax
'0070c13f    e95cfeffff              jmp 70bfa0

'ERROR: Two many next close:
Loop
'0070c144    b801000000              mov eax, 00000001
'0070c149    660345e8                add ax, word ptr [ebp-18]
var_num1 = 1 + var_41
'0070c14d    0f80430b0000            jo 70cc96
'0070c153    8945e8                  mov dword ptr [ebp-18], eax
'0070c156    e9dffdffff              jmp 70bf3a

'ERROR: Two many next close:
Loop
'0070c15b    8b16                    mov edx, dword ptr [esi]
'0070c15d    56                      push esi
'0070c15e    ff9220030000            call dword ptr [edx+00000320]
'0070c164    50                      push eax
'0070c165    8d45b8                  lea eax, dword ptr [ebp-48]
'0070c168    50                      push eax
'0070c169    ffd3                    call ebx
Set var_61 = CInt(var_num3)
'0070c16b    50                      push eax
'0070c16c    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0070c16f    51                      push ecx
'0070c170    ffd7                    call edi
var_121 = var_61.UNK__6
'0070c172    83c410                  add esp, 10
'0070c175    50                      push eax

' *** Reference to "__vbaI4Var"
'0070c176    ff157c124000            call dword ptr [0040127c]
'0070c17c    8bc8                    mov ecx, eax
'0070c17e    83e901                  sub ecx, 01
var_num3 = CLng(var_121) - 1
'0070c181    0f800f0b0000            jo 70cc96

' *** Reference to "__vbaI2I4"
'0070c187    ff1550114000            call dword ptr [00401150]
'0070c18d    8d4db8                  lea ecx, dword ptr [ebp-48]
'0070c190    c745e801000000          mov dword ptr [ebp-18], 00000001
'0070c197    8985dcfdffff            mov dword ptr [ebp+fffffddc], eax

' *** Reference to "__vbaFreeObj"
'0070c19d    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'0070c1a3    8d4d94                  lea ecx, dword ptr [ebp-6c]

' *** Reference to "__vbaFreeVar"
'0070c1a6    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_121)
'0070c1ac    668b95dcfdffff          mov dx, word ptr [ebp+fffffddc]
'0070c1b3    663955e8                cmp word ptr [ebp-18], dx
'0070c1b7    0f8f53090000            jg 70cb10

If (var_41 <= CInt(var_num3)) Then
'0070c1bd    8b06                    mov eax, dword ptr [esi]
'0070c1bf    6a00                    push 00
'0070c1c1    6a07                    push 07
'0070c1c3    56                      push esi
'0070c1c4    c745e000000000          mov dword ptr [ebp-20], 00000000
'0070c1cb    ff9020030000            call dword ptr [eax+00000320]
'0070c1d1    50                      push eax
'0070c1d2    8d4db8                  lea ecx, dword ptr [ebp-48]
'0070c1d5    51                      push ecx
'0070c1d6    ffd3                    call ebx
    Set var_61 = Nothing
'0070c1d8    50                      push eax
'0070c1d9    8d5594                  lea edx, dword ptr [ebp-6c]
'0070c1dc    52                      push edx
'0070c1dd    ffd7                    call edi
    var_121 = var_61.UNK_-256 - 12_7
'0070c1df    83c410                  add esp, 10
'0070c1e2    50                      push eax

' *** Reference to "__vbaI4Var"
'0070c1e3    ff157c124000            call dword ptr [0040127c]
'0070c1e9    8bc8                    mov ecx, eax
'0070c1eb    83e901                  sub ecx, 01
    var_num3 = CLng(var_121) - 1
'0070c1ee    0f80a20a0000            jo 70cc96

' *** Reference to "__vbaI2I4"
'0070c1f4    ff1550114000            call dword ptr [00401150]
'0070c1fa    8d4db8                  lea ecx, dword ptr [ebp-48]
'0070c1fd    c745e400000000          mov dword ptr [ebp-1c], 00000000
'0070c204    8985d4fdffff            mov dword ptr [ebp+fffffdd4], eax

' *** Reference to "__vbaFreeObj"
'0070c20a    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'0070c210    8d4d94                  lea ecx, dword ptr [ebp-6c]

' *** Reference to "__vbaFreeVar"
'0070c213    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_121)
'0070c219    8b45e4                  mov eax, dword ptr [ebp-1c]
'0070c21c    663b85d4fdffff          cmp ax, word ptr [ebp+fffffdd4]
'0070c223    0f8fa2000000            jg 70c2cb
    
    Do While (    var_40 <= WORD PTR [EBP+FFFFFDD4])
'0070c229    83ec10                  sub esp, 10
'0070c22c    8bd4                    mov edx, esp
'0070c22e    b903000000              mov ecx, 00000003
'0070c233    890a                    mov dword ptr [edx], ecx
'0070c235    0fbfc0                  movsx eax, ax
'0070c238    898d14ffffff            mov dword ptr [ebp+ffffff14], ecx
'0070c23e    8b8d18ffffff            mov ecx, dword ptr [ebp+ffffff18]
'0070c244    894a04                  mov dword ptr [edx+04], ecx
'0070c247    8b0e                    mov ecx, dword ptr [esi]
'0070c249    6a01                    push 01
'0070c24b    894208                  mov dword ptr [edx+08], eax
'0070c24e    89851cffffff            mov dword ptr [ebp+ffffff1c], eax
'0070c254    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'0070c25a    68a8000000              push 000000a8
'0070c25f    56                      push esi
'0070c260    89420c                  mov dword ptr [edx+0c], eax
'0070c263    ff9120030000            call dword ptr [ecx+00000320]
'0070c269    50                      push eax
'0070c26a    8d55b8                  lea edx, dword ptr [ebp-48]
'0070c26d    52                      push edx
'0070c26e    ffd3                    call ebx
    Set var_61 = Nothing
'0070c270    50                      push eax
'0070c271    8d4594                  lea eax, dword ptr [ebp-6c]
'0070c274    50                      push eax
'0070c275    ffd7                    call edi
    var_121 = var_61.UNK_-256 - 12_168
'0070c277    83c420                  add esp, 20
'0070c27a    50                      push eax

' *** Reference to "__vbaI4Var"
'0070c27b    ff157c124000            call dword ptr [0040127c]
'0070c281    0fbf4de8                movsx ecx, word ptr [ebp-18]
'0070c285    8bf8                    mov edi, eax
'0070c287    2bf9                    sub edi, ecx
    var_num7 = CLng(var_121) - var_41
'0070c289    f7df                    neg edi
'0070c28b    1bff                    sbb edi, edi
'0070c28d    47                      inc edi
'0070c28e    8d4db8                  lea ecx, dword ptr [ebp-48]
'0070c291    f7df                    neg edi

' *** Reference to "__vbaFreeObj"
'0070c293    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_61)
'0070c299    8d4d94                  lea ecx, dword ptr [ebp-6c]

' *** Reference to "__vbaFreeVar"
'0070c29c    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_121)
'0070c2a2    6685ff                  test edi, edi
'0070c2a5    7407                    je 70c2ae
    
    If (    -(CBool((var_num7)))) Then
'0070c2a7    c745e0ffffffff          mov dword ptr [ebp-20], ffffffff
    
End If

' *** Reference to "__vbaLateIdCallLd"
'0070c2ae    8b3d7c114000            mov edi, dword ptr [0040117c]
'0070c2b4    b801000000              mov eax, 00000001
'0070c2b9    660345e4                add ax, word ptr [ebp-1c]
var_num1 = 1 + var_40
'0070c2bd    0f80d3090000            jo 70cc96
'0070c2c3    8945e4                  mov dword ptr [ebp-1c], eax
'0070c2c6    e94effffff              jmp 70c219

'ERROR: Two many next close:
Loop
'0070c2cb    66837de000              cmp word ptr [ebp-20], 00
'0070c2d0    0f8499070000            je 70ca6f

If (var_3 <> 0) Then
'0070c2d6    8b16                    mov edx, dword ptr [esi]
'0070c2d8    56                      push esi
'0070c2d9    ff9204030000            call dword ptr [edx+00000304]
'0070c2df    50                      push eax
'0070c2e0    8d45b8                  lea eax, dword ptr [ebp-48]
'0070c2e3    50                      push eax
'0070c2e4    ffd3                    call ebx
    Set var_61 = Nothing
'0070c2e6    8d55d0                  lea edx, dword ptr [ebp-30]
'0070c2e9    8bf8                    mov edi, eax
'0070c2eb    8b0f                    mov ecx, dword ptr [edi]
'0070c2ed    52                      push edx
'0070c2ee    57                      push edi
'0070c2ef    ff91a0000000            call dword ptr [ecx+000000a0]
'0070c2f5    dbe2                    fnclex
'0070c2f7    85c0                    test ax, ax
'0070c2f9    7d12                    jge 70c30d
'0070c2fb    68a0000000              push 000000a0
'0070c300    68d00d4300              push 00430dd0
'0070c305    57                      push edi
'0070c306    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070c307    ff1580104000            call dword ptr [00401080]
'0070c30d    8b45d0                  mov eax, dword ptr [ebp-30]
'0070c310    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'0070c311    ff1510134000            call dword ptr [00401310]
'0070c317    dd9d18feffff            fstp qword ptr [ebp+fffffe18]
    'var_694 = (00)
'0070c31d    8b0e                    mov ecx, dword ptr [esi]
'0070c31f    56                      push esi
'0070c320    ff9104030000            call dword ptr [ecx+00000304]
'0070c326    50                      push eax
'0070c327    8d55b4                  lea edx, dword ptr [ebp-4c]
'0070c32a    52                      push edx
'0070c32b    ffd3                    call ebx
    Set var_62 = Nothing
'0070c32d    8b8d18feffff            mov ecx, dword ptr [ebp+fffffe18]
'0070c333    8b951cfeffff            mov edx, dword ptr [ebp+fffffe1c]
'0070c339    894d9c                  mov dword ptr [ebp-64], ecx
'0070c33c    c7458c7d000000          mov dword ptr [ebp-74], 0000007d
'0070c343    c7458402000000          mov dword ptr [ebp-7c], 00000002
'0070c34a    8955a0                  mov dword ptr [ebp-60], edx
'0070c34d    c7459405000000          mov dword ptr [ebp-6c], 00000005
'0070c354    8b38                    mov edi, dword ptr [eax]
'0070c356    89850cfeffff            mov dword ptr [ebp+fffffe0c], eax
'0070c35c    8d4584                  lea eax, dword ptr [ebp-7c]
'0070c35f    50                      push eax
'0070c360    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0070c363    51                      push ecx
'0070c364    e82764deff              call 4f2790
    Call sub_4F2790()
'0070c369    50                      push eax

' *** Reference to "__vbaStrI2"
'0070c36a    ff1510104000            call dword ptr [00401010]
'0070c370    8bd0                    mov edx, eax
'0070c372    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaStrMove"
'0070c375    ff15d0124000            call dword ptr [004012d0]
'0070c37b    8bd7                    mov edx, edi
'0070c37d    8bbd0cfeffff            mov edi, dword ptr [ebp+fffffe0c]
'0070c383    50                      push eax
'0070c384    57                      push edi
'0070c385    ff92a4000000            call dword ptr [edx+000000a4]
'0070c38b    dbe2                    fnclex
'0070c38d    85c0                    test ax, ax
'0070c38f    7d12                    jge 70c3a3
'0070c391    68a4000000              push 000000a4
'0070c396    68d00d4300              push 00430dd0
'0070c39b    57                      push edi
'0070c39c    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070c39d    ff1580104000            call dword ptr [00401080]
'0070c3a3    8d45cc                  lea eax, dword ptr [ebp-34]
'0070c3a6    50                      push eax
'0070c3a7    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0070c3aa    51                      push ecx
'0070c3ab    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'0070c3ad    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( , -4748)
'0070c3b3    8d55b4                  lea edx, dword ptr [ebp-4c]
'0070c3b6    52                      push edx
'0070c3b7    8d45b8                  lea eax, dword ptr [ebp-48]
'0070c3ba    50                      push eax
'0070c3bb    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070c3bd    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_61, var_62)
'0070c3c3    8d4d84                  lea ecx, dword ptr [ebp-7c]
'0070c3c6    51                      push ecx
'0070c3c7    8d5594                  lea edx, dword ptr [ebp-6c]
'0070c3ca    52                      push edx
'0070c3cb    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'0070c3cd    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_121, var_119)
'0070c3d3    8b06                    mov eax, dword ptr [esi]
'0070c3d5    83c424                  add esp, 24
'0070c3d8    56                      push esi
'0070c3d9    ff9004030000            call dword ptr [eax+00000304]
'0070c3df    50                      push eax
'0070c3e0    8d4db8                  lea ecx, dword ptr [ebp-48]
'0070c3e3    51                      push ecx
'0070c3e4    ffd3                    call ebx
    Set var_61 = Nothing
'0070c3e6    8bf8                    mov edi, eax
'0070c3e8    8b17                    mov edx, dword ptr [edi]
'0070c3ea    8d45d0                  lea eax, dword ptr [ebp-30]
'0070c3ed    50                      push eax
'0070c3ee    57                      push edi
'0070c3ef    ff92a0000000            call dword ptr [edx+000000a0]
'0070c3f5    dbe2                    fnclex
'0070c3f7    85c0                    test ax, ax
'0070c3f9    7d12                    jge 70c40d
'0070c3fb    68a0000000              push 000000a0
'0070c400    68d00d4300              push 00430dd0
'0070c405    57                      push edi
'0070c406    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070c407    ff1580104000            call dword ptr [00401080]
'0070c40d    668b55e8                mov dx, word ptr [ebp-18]
'0070c411    83ec10                  sub esp, 10
'0070c414    8bfc                    mov edi, esp
'0070c416    b903000000              mov ecx, 00000003
'0070c41b    890f                    mov dword ptr [edi], ecx
'0070c41d    898d14ffffff            mov dword ptr [ebp+ffffff14], ecx
'0070c423    8b8d18ffffff            mov ecx, dword ptr [ebp+ffffff18]
'0070c429    894f04                  mov dword ptr [edi+04], ecx
'0070c42c    33c0                    xor eax, eax
'0070c42e    894708                  mov dword ptr [edi+08], eax
'0070c431    83ec10                  sub esp, 10
'0070c434    89851cffffff            mov dword ptr [ebp+ffffff1c], eax
'0070c43a    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'0070c440    89470c                  mov dword ptr [edi+0c], eax
'0070c443    668995fcfeffff          mov word ptr [ebp+fffffefc], dx
'0070c44a    8b85fcfeffff            mov eax, dword ptr [ebp+fffffefc]
'0070c450    8bcc                    mov ecx, esp
'0070c452    ba02000000              mov edx, 00000002
'0070c457    8911                    mov dword ptr [ecx], edx
'0070c459    8b95f8feffff            mov edx, dword ptr [ebp+fffffef8]
'0070c45f    895104                  mov dword ptr [ecx+04], edx
'0070c462    8b9500ffffff            mov edx, dword ptr [ebp+ffffff00]
'0070c468    894108                  mov dword ptr [ecx+08], eax
'0070c46b    89510c                  mov dword ptr [ecx+0c], edx
'0070c46e    8b95d8feffff            mov edx, dword ptr [ebp+fffffed8]
'0070c474    83ec10                  sub esp, 10
'0070c477    8bcc                    mov ecx, esp
'0070c479    b802000000              mov eax, 00000002
'0070c47e    8901                    mov dword ptr [ecx], eax
'0070c480    895104                  mov dword ptr [ecx+04], edx
'0070c483    894108                  mov dword ptr [ecx+08], eax
'0070c486    8b85e0feffff            mov eax, dword ptr [ebp+fffffee0]
'0070c48c    bf03000000              mov edi, 00000003
'0070c491    57                      push edi
'0070c492    689e000000              push 0000009e
'0070c497    89410c                  mov dword ptr [ecx+0c], eax
'0070c49a    8b0e                    mov ecx, dword ptr [esi]
'0070c49c    56                      push esi
'0070c49d    ff9120030000            call dword ptr [ecx+00000320]
'0070c4a3    50                      push eax
'0070c4a4    8d55b4                  lea edx, dword ptr [ebp-4c]
'0070c4a7    52                      push edx
'0070c4a8    ffd3                    call ebx
    Set var_62 = Nothing
'0070c4aa    50                      push eax
'0070c4ab    8d4594                  lea eax, dword ptr [ebp-6c]
'0070c4ae    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'0070c4af    ff157c114000            call dword ptr [0040117c]
    var_121 = var_62.UNK_-256 - 12_158
'0070c4b5    83c440                  add esp, 40
'0070c4b8    50                      push eax
'0070c4b9    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0070c4bc    51                      push ecx

' *** Reference to "__vbaStrVarVal"
'0070c4bd    ff15fc114000            call dword ptr [004011fc]
'0070c4c3    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'0070c4c4    ff1510134000            call dword ptr [00401310]
'0070c4ca    dd9d18feffff            fstp qword ptr [ebp+fffffe18]
    'var_694 = (00)
'0070c4d0    8b9518feffff            mov edx, dword ptr [ebp+fffffe18]
'0070c4d6    8b851cfeffff            mov eax, dword ptr [ebp+fffffe1c]
'0070c4dc    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'0070c4e2    89558c                  mov dword ptr [ebp-74], edx
'0070c4e5    51                      push ecx
'0070c4e6    8d5584                  lea edx, dword ptr [ebp-7c]
'0070c4e9    52                      push edx
'0070c4ea    89bd7cffffff            mov dword ptr [ebp+ffffff7c], edi
'0070c4f0    c78574ffffff02000000    mov dword ptr [ebp+ffffff74], 00000002
'0070c4fa    894590                  mov dword ptr [ebp-70], eax
'0070c4fd    c7458405000000          mov dword ptr [ebp-7c], 00000005
'0070c504    e8e761deff              call 4f26f0
    Call sub_4F26F0()
'0070c509    8b154cb07200            mov edx, dword ptr [0072b04c]
'0070c50f    c7859cfeffff02000000    mov dword ptr [ebp+fffffe9c], 00000002
'0070c519    89bd94feffff            mov dword ptr [ebp+fffffe94], edi
'0070c51f    8b3a                    mov edi, dword ptr [edx]
'0070c521    8d55b0                  lea edx, dword ptr [ebp-50]
'0070c524    52                      push edx
'0070c525    83ec10                  sub esp, 10
'0070c528    898520feffff            mov dword ptr [ebp+fffffe20], eax
'0070c52e    b90a000000              mov ecx, 0000000a
'0070c533    b804000280              mov eax, 80020004
'0070c538    8bd4                    mov edx, esp
'0070c53a    89858cfeffff            mov dword ptr [ebp+fffffe8c], eax
'0070c540    898d84feffff            mov dword ptr [ebp+fffffe84], ecx
'0070c546    890a                    mov dword ptr [edx], ecx
'0070c548    8b8d78feffff            mov ecx, dword ptr [ebp+fffffe78]
'0070c54e    894a04                  mov dword ptr [edx+04], ecx
'0070c551    894208                  mov dword ptr [edx+08], eax
'0070c554    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'0070c55a    89420c                  mov dword ptr [edx+0c], eax
'0070c55d    8b9584feffff            mov edx, dword ptr [ebp+fffffe84]
'0070c563    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'0070c569    83ec10                  sub esp, 10
'0070c56c    8bcc                    mov ecx, esp
'0070c56e    8911                    mov dword ptr [ecx], edx
'0070c570    8b958cfeffff            mov edx, dword ptr [ebp+fffffe8c]
'0070c576    894104                  mov dword ptr [ecx+04], eax
'0070c579    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'0070c57f    895108                  mov dword ptr [ecx+08], edx
'0070c582    8b9594feffff            mov edx, dword ptr [ebp+fffffe94]
'0070c588    89410c                  mov dword ptr [ecx+0c], eax
'0070c58b    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0070c591    83ec10                  sub esp, 10
'0070c594    8bcc                    mov ecx, esp
'0070c596    8911                    mov dword ptr [ecx], edx
'0070c598    8b959cfeffff            mov edx, dword ptr [ebp+fffffe9c]
'0070c59e    894104                  mov dword ptr [ecx+04], eax
'0070c5a1    8b85a0feffff            mov eax, dword ptr [ebp+fffffea0]
'0070c5a7    895108                  mov dword ptr [ecx+08], edx
'0070c5aa    89410c                  mov dword ptr [ecx+0c], eax
'0070c5ad    8b4dd0                  mov ecx, dword ptr [ebp-30]
'0070c5b0    6838644500              push 00456438
'0070c5b5    51                      push ecx

' *** Reference to "__vbaStrCat"
'0070c5b6    ff1570104000            call dword ptr [00401070]
    var_49 = ("select fp") & (vbNullString)
'0070c5bc    8bd0                    mov edx, eax
'0070c5be    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaStrMove"
'0070c5c1    ff15d0124000            call dword ptr [004012d0]
'0070c5c7    50                      push eax
'0070c5c8    6850644500              push 00456450

' *** Reference to "__vbaStrCat"
'0070c5cd    ff1570104000            call dword ptr [00401070]
    var_2153 = (var_49) & (" from experience where niveau=")
'0070c5d3    8bd0                    mov edx, eax
'0070c5d5    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaStrMove"
'0070c5d8    ff15d0124000            call dword ptr [004012d0]
'0070c5de    8b9520feffff            mov edx, dword ptr [ebp+fffffe20]
'0070c5e4    50                      push eax
'0070c5e5    52                      push edx

' *** Reference to "__vbaStrI2"
'0070c5e6    ff1510104000            call dword ptr [00401010]
'0070c5ec    8bd0                    mov edx, eax
'0070c5ee    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaStrMove"
'0070c5f1    ff15d0124000            call dword ptr [004012d0]
'0070c5f7    50                      push eax

' *** Reference to "__vbaStrCat"
'0070c5f8    ff1570104000            call dword ptr [00401070]
    var_1532 = (var_2153) & (CStr(VB[CODE]))
'0070c5fe    8bd0                    mov edx, eax
'0070c600    8d4dbc                  lea ecx, dword ptr [ebp-44]

' *** Reference to "__vbaStrMove"
'0070c603    ff15d0124000            call dword ptr [004012d0]
'0070c609    50                      push eax
'0070c60a    a14cb07200              mov ax, word ptr [0072b04c]
'0070c60f    50                      push eax
'0070c610    ff97bc000000            call dword ptr [edi+000000bc]
'0070c616    dbe2                    fnclex
'0070c618    85c0                    test ax, ax
'0070c61a    7d18                    jge 70c634
'0070c61c    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'0070c622    68bc000000              push 000000bc
'0070c627    68301f4300              push 00431f30
'0070c62c    51                      push ecx
'0070c62d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070c62e    ff1580104000            call dword ptr [00401080]
'0070c634    8b45b0                  mov eax, dword ptr [ebp-50]
'0070c637    50                      push eax
'0070c638    8d55d4                  lea edx, dword ptr [ebp-2c]
'0070c63b    52                      push edx
'0070c63c    c745b000000000          mov dword ptr [ebp-50], 00000000
'0070c643    ffd3                    call ebx
    Set var_44 = Nothing
'0070c645    8d45bc                  lea eax, dword ptr [ebp-44]
'0070c648    50                      push eax
'0070c649    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0070c64c    51                      push ecx
'0070c64d    8d55c4                  lea edx, dword ptr [ebp-3c]
'0070c650    52                      push edx
'0070c651    8d45c8                  lea eax, dword ptr [ebp-38]
'0070c654    50                      push eax
'0070c655    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070c658    51                      push ecx
'0070c659    8d55d0                  lea edx, dword ptr [ebp-30]
'0070c65c    52                      push edx
'0070c65d    6a06                    push 06

' *** Reference to "__vbaFreeStrList"
'0070c65f    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( , -4780, 0, -4784, -4788, -4792)
'0070c665    8d45b4                  lea eax, dword ptr [ebp-4c]
'0070c668    50                      push eax
'0070c669    8d4db8                  lea ecx, dword ptr [ebp-48]
'0070c66c    51                      push ecx
'0070c66d    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070c66f    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_61, var_62)
'0070c675    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'0070c67b    52                      push edx
'0070c67c    8d4584                  lea eax, dword ptr [ebp-7c]
'0070c67f    50                      push eax
'0070c680    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0070c683    51                      push ecx
'0070c684    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'0070c686    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_121, var_119, var_91)
'0070c68c    8b16                    mov edx, dword ptr [esi]
'0070c68e    83c438                  add esp, 38
'0070c691    56                      push esi
'0070c692    ff9204030000            call dword ptr [edx+00000304]
'0070c698    50                      push eax
'0070c699    8d45b4                  lea eax, dword ptr [ebp-4c]
'0070c69c    50                      push eax
'0070c69d    ffd3                    call ebx
    Set var_62 = 
'0070c69f    8d55cc                  lea edx, dword ptr [ebp-34]
'0070c6a2    8bf8                    mov edi, eax
'0070c6a4    8b0f                    mov ecx, dword ptr [edi]
'0070c6a6    52                      push edx
'0070c6a7    57                      push edi
'0070c6a8    ff91a0000000            call dword ptr [ecx+000000a0]
'0070c6ae    dbe2                    fnclex
'0070c6b0    85c0                    test ax, ax
'0070c6b2    7d12                    jge 70c6c6
'0070c6b4    68a0000000              push 000000a0
'0070c6b9    68d00d4300              push 00430dd0
'0070c6be    57                      push edi
'0070c6bf    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070c6c0    ff1580104000            call dword ptr [00401080]
'0070c6c6    8b06                    mov eax, dword ptr [esi]
'0070c6c8    56                      push esi
'0070c6c9    ff90fc020000            call dword ptr [eax+000002fc]
'0070c6cf    50                      push eax
'0070c6d0    8d4da8                  lea ecx, dword ptr [ebp-58]
'0070c6d3    51                      push ecx
'0070c6d4    ffd3                    call ebx
    Set var_86 = Nothing
'0070c6d6    8bf8                    mov edi, eax
'0070c6d8    8b17                    mov edx, dword ptr [edi]
'0070c6da    8d45c8                  lea eax, dword ptr [ebp-38]
'0070c6dd    50                      push eax
'0070c6de    57                      push edi
'0070c6df    ff92a0000000            call dword ptr [edx+000000a0]
'0070c6e5    dbe2                    fnclex
'0070c6e7    85c0                    test ax, ax
'0070c6e9    7d12                    jge 70c6fd
'0070c6eb    68a0000000              push 000000a0
'0070c6f0    68d00d4300              push 00430dd0
'0070c6f5    57                      push edi
'0070c6f6    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070c6f7    ff1580104000            call dword ptr [00401080]
'0070c6fd    8b4dc8                  mov ecx, dword ptr [ebp-38]
'0070c700    51                      push ecx

' *** Reference to "rtcR8ValFromBstr"
'0070c701    ff1510134000            call dword ptr [00401310]
'0070c707    dd9d18feffff            fstp qword ptr [ebp+fffffe18]
    'var_694 = (00)
'0070c70d    8b55e8                  mov edx, dword ptr [ebp-18]
'0070c710    83ec10                  sub esp, 10
'0070c713    8bfc                    mov edi, esp
'0070c715    33c0                    xor eax, eax
'0070c717    b903000000              mov ecx, 00000003
'0070c71c    890f                    mov dword ptr [edi], ecx
'0070c71e    898d14ffffff            mov dword ptr [ebp+ffffff14], ecx
'0070c724    8b8d18ffffff            mov ecx, dword ptr [ebp+ffffff18]
'0070c72a    894f04                  mov dword ptr [edi+04], ecx
'0070c72d    894708                  mov dword ptr [edi+08], eax
'0070c730    89851cffffff            mov dword ptr [ebp+ffffff1c], eax
'0070c736    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'0070c73c    89470c                  mov dword ptr [edi+0c], eax
'0070c73f    83ec10                  sub esp, 10
'0070c742    668995fcfeffff          mov word ptr [ebp+fffffefc], dx
'0070c749    8b85fcfeffff            mov eax, dword ptr [ebp+fffffefc]
'0070c74f    6689956cfeffff          mov word ptr [ebp+fffffe6c], dx
'0070c756    8bcc                    mov ecx, esp
'0070c758    ba02000000              mov edx, 00000002
'0070c75d    8911                    mov dword ptr [ecx], edx
'0070c75f    8b95f8feffff            mov edx, dword ptr [ebp+fffffef8]
'0070c765    895104                  mov dword ptr [ecx+04], edx
'0070c768    8b9500ffffff            mov edx, dword ptr [ebp+ffffff00]
'0070c76e    894108                  mov dword ptr [ecx+08], eax
'0070c771    89510c                  mov dword ptr [ecx+0c], edx
'0070c774    8b95d8feffff            mov edx, dword ptr [ebp+fffffed8]
'0070c77a    83ec10                  sub esp, 10
'0070c77d    8bcc                    mov ecx, esp
'0070c77f    b802000000              mov eax, 00000002
'0070c784    8901                    mov dword ptr [ecx], eax
'0070c786    895104                  mov dword ptr [ecx+04], edx
'0070c789    b806000000              mov eax, 00000006
'0070c78e    894108                  mov dword ptr [ecx+08], eax
'0070c791    8b85e0feffff            mov eax, dword ptr [ebp+fffffee0]
'0070c797    6a03                    push 03
'0070c799    689e000000              push 0000009e
'0070c79e    89410c                  mov dword ptr [ecx+0c], eax
'0070c7a1    8b0e                    mov ecx, dword ptr [esi]
'0070c7a3    56                      push esi
'0070c7a4    ff9120030000            call dword ptr [ecx+00000320]
'0070c7aa    50                      push eax
'0070c7ab    8d55b8                  lea edx, dword ptr [ebp-48]
'0070c7ae    52                      push edx
'0070c7af    ffd3                    call ebx
    Set var_61 = Nothing
'0070c7b1    50                      push eax
'0070c7b2    8d4594                  lea eax, dword ptr [ebp-6c]
'0070c7b5    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'0070c7b6    ff157c114000            call dword ptr [0040117c]
    var_121 = var_61.UNK_-256 - 12_158
'0070c7bc    83c440                  add esp, 40
'0070c7bf    50                      push eax
'0070c7c0    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0070c7c3    51                      push ecx

' *** Reference to "__vbaStrVarVal"
'0070c7c4    ff15fc114000            call dword ptr [004011fc]
'0070c7ca    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'0070c7cb    ff1510134000            call dword ptr [00401310]
'0070c7d1    dd9dbcfeffff            fstp qword ptr [ebp+fffffebc]
    'var_482 = (00)
'0070c7d7    8b45d4                  mov eax, dword ptr [ebp-2c]
'0070c7da    8d4db0                  lea ecx, dword ptr [ebp-50]
'0070c7dd    51                      push ecx
'0070c7de    c785b4feffff05000000    mov dword ptr [ebp+fffffeb4], 00000005
'0070c7e8    8b10                    mov edx, dword ptr [eax]
'0070c7ea    50                      push eax
'0070c7eb    ff92b4000000            call dword ptr [edx+000000b4]
'0070c7f1    dbe2                    fnclex
'0070c7f3    85c0                    test ax, ax
'0070c7f5    7d15                    jge 70c80c
'0070c7f7    8b55d4                  mov edx, dword ptr [ebp-2c]
'0070c7fa    68b4000000              push 000000b4
'0070c7ff    6830314300              push 00433130
'0070c804    52                      push edx
'0070c805    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070c806    ff1580104000            call dword ptr [00401080]
'0070c80c    8b45cc                  mov eax, dword ptr [ebp-34]
'0070c80f    8b7db0                  mov edi, dword ptr [ebp-50]
'0070c812    6894644500              push 00456494
'0070c817    50                      push eax

' *** Reference to "__vbaStrCat"
'0070c818    ff1570104000            call dword ptr [00401070]
    var_2153 = ("fp") & (var_49)
'0070c81e    8d55ac                  lea edx, dword ptr [ebp-54]
'0070c821    52                      push edx
'0070c822    89458c                  mov dword ptr [ebp-74], eax
'0070c825    83ec10                  sub esp, 10
'0070c828    b808000000              mov eax, 00000008
'0070c82d    894584                  mov dword ptr [ebp-7c], eax
'0070c830    8b0f                    mov ecx, dword ptr [edi]
'0070c832    8bd4                    mov edx, esp
'0070c834    8902                    mov dword ptr [edx], eax
'0070c836    8b4588                  mov eax, dword ptr [ebp-78]
'0070c839    894204                  mov dword ptr [edx+04], eax
'0070c83c    8b458c                  mov eax, dword ptr [ebp-74]
'0070c83f    894208                  mov dword ptr [edx+08], eax
'0070c842    8b4590                  mov eax, dword ptr [ebp-70]
'0070c845    57                      push edi
'0070c846    89420c                  mov dword ptr [edx+0c], eax
'0070c849    ff5130                  call dword ptr [ecx+30]
'0070c84c    dbe2                    fnclex
'0070c84e    85c0                    test ax, ax
'0070c850    7d0f                    jge 70c861
'0070c852    6a30                    push 30
'0070c854    68d8304300              push 004330d8
'0070c859    57                      push edi
'0070c85a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070c85b    ff1580104000            call dword ptr [00401080]
'0070c861    8b45ac                  mov eax, dword ptr [ebp-54]
'0070c864    8b08                    mov ecx, dword ptr [eax]
'0070c866    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'0070c86c    52                      push edx
'0070c86d    50                      push eax
'0070c86e    8bf8                    mov edi, eax
'0070c870    ff5144                  call dword ptr [ecx+44]
'0070c873    dbe2                    fnclex
'0070c875    85c0                    test ax, ax
'0070c877    7d0f                    jge 70c888
'0070c879    6a44                    push 44
'0070c87b    6880304300              push 00433080
'0070c880    57                      push edi
'0070c881    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070c882    ff1580104000            call dword ptr [00401080]
'0070c888    dd8518feffff            fld qword ptr [ebp+fffffe18]
'0070c88e    8b45d8                  mov eax, dword ptr [ebp-28]
'0070c891    833d00b0720000          cmp dword ptr [0072b000], 00
'0070c898    7508                    jne 70c8a2
    
    If (    vbNullChar = 0) Then
'0070c89a    dc35286c4000            fdiv qword ptr [00406c28]
'0070c8a0    eb11                    jmp 70c8b3
    
End If
'0070c8a2    ff352c6c4000            push dword ptr [00406c2c]
'0070c8a8    ff35286c4000            push dword ptr [00406c28]

' *** Reference to "_adj_fdiv_m64"
'0070c8ae    e8d1a9cfff              call 407284
'0070c8b3    8b4ddc                  mov ecx, dword ptr [ebp-24]
'0070c8b6    8985acfeffff            mov dword ptr [ebp+fffffeac], eax
'0070c8bc    898db0feffff            mov dword ptr [ebp+fffffeb0], ecx
'0070c8c2    b905000000              mov ecx, 00000005
'0070c8c7    898da4feffff            mov dword ptr [ebp+fffffea4], ecx
'0070c8cd    898d94feffff            mov dword ptr [ebp+fffffe94], ecx
'0070c8d3    8b8d90feffff            mov ecx, dword ptr [ebp+fffffe90]
'0070c8d9    dc2d58164000            fsubr qword ptr [00401658]
'0070c8df    dd9d9cfeffff            fstp qword ptr [ebp+fffffe9c]
'var_822 = (00)
'0070c8e5    dfe0                    fnstsw ax
'0070c8e7    a80d                    test al, 0d
'0070c8e9    0f85a2030000            jne 70cc91
'0070c8ef    83ec10                  sub esp, 10
'0070c8f2    8bd4                    mov edx, esp
'0070c8f4    b803000000              mov eax, 00000003
'0070c8f9    8902                    mov dword ptr [edx], eax
'0070c8fb    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'0070c901    894204                  mov dword ptr [edx+04], eax
'0070c904    33c0                    xor eax, eax
var_num1 = Empty
'0070c906    894208                  mov dword ptr [edx+08], eax
'0070c909    894a0c                  mov dword ptr [edx+0c], ecx
'0070c90c    8b8d6cfeffff            mov ecx, dword ptr [ebp+fffffe6c]
'0070c912    83ec10                  sub esp, 10
'0070c915    8bd4                    mov edx, esp
'0070c917    b802000000              mov eax, 00000002
'0070c91c    8902                    mov dword ptr [edx], eax
'0070c91e    8b8568feffff            mov eax, dword ptr [ebp+fffffe68]
'0070c924    894204                  mov dword ptr [edx+04], eax
'0070c927    8b8570feffff            mov eax, dword ptr [ebp+fffffe70]
'0070c92d    894a08                  mov dword ptr [edx+08], ecx
'0070c930    89420c                  mov dword ptr [edx+0c], eax
'0070c933    8b9548feffff            mov edx, dword ptr [ebp+fffffe48]
'0070c939    83ec10                  sub esp, 10
'0070c93c    8bcc                    mov ecx, esp
'0070c93e    b802000000              mov eax, 00000002
'0070c943    8901                    mov dword ptr [ecx], eax
'0070c945    895104                  mov dword ptr [ecx+04], edx
'0070c948    b805000000              mov eax, 00000005
'0070c94d    894108                  mov dword ptr [ecx+08], eax
'0070c950    8b8550feffff            mov eax, dword ptr [ebp+fffffe50]
'0070c956    89410c                  mov dword ptr [ecx+0c], eax
'0070c959    8d8db4feffff            lea ecx, dword ptr [ebp+fffffeb4]
'0070c95f    51                      push ecx
'0070c960    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'0070c966    52                      push edx
'0070c967    8d8564ffffff            lea eax, dword ptr [ebp+ffffff64]
'0070c96d    50                      push eax

' *** Reference to "__vbaVarMul"
'0070c96e    ff15a8114000            call dword ptr [004011a8]
'0070c974    50                      push eax
'0070c975    8d8da4feffff            lea ecx, dword ptr [ebp+fffffea4]
'0070c97b    51                      push ecx
'0070c97c    8d9554ffffff            lea edx, dword ptr [ebp+ffffff54]
'0070c982    52                      push edx

' *** Reference to "__vbaVarDiv"
'0070c983    ff15d8114000            call dword ptr [004011d8]

' *** Reference to "__vbaVarInt"
'0070c989    8b3d38124000            mov edi, dword ptr [00401238]
'0070c98f    50                      push eax
'0070c990    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'0070c996    50                      push eax
'0070c997    ffd7                    call edi
'0070c999    50                      push eax
'0070c99a    8d8d94feffff            lea ecx, dword ptr [ebp+fffffe94]
'0070c9a0    51                      push ecx
'0070c9a1    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'0070c9a7    52                      push edx

' *** Reference to "__vbaVarMul"
'0070c9a8    ff15a8114000            call dword ptr [004011a8]
'0070c9ae    50                      push eax
'0070c9af    8d8524ffffff            lea eax, dword ptr [ebp+ffffff24]
'0070c9b5    50                      push eax
'0070c9b6    ffd7                    call edi
'0070c9b8    8b10                    mov edx, dword ptr [eax]
'0070c9ba    83ec10                  sub esp, 10
'0070c9bd    8bcc                    mov ecx, esp
'0070c9bf    8911                    mov dword ptr [ecx], edx
'0070c9c1    8b5004                  mov edx, dword ptr [eax+04]
'0070c9c4    895104                  mov dword ptr [ecx+04], edx
'0070c9c7    8b5008                  mov edx, dword ptr [eax+08]
'0070c9ca    895108                  mov dword ptr [ecx+08], edx
'0070c9cd    8b400c                  mov eax, dword ptr [eax+0c]
'0070c9d0    6a03                    push 03
'0070c9d2    689e000000              push 0000009e
'0070c9d7    89410c                  mov dword ptr [ecx+0c], eax
'0070c9da    8b0e                    mov ecx, dword ptr [esi]
'0070c9dc    56                      push esi
'0070c9dd    ff9120030000            call dword ptr [ecx+00000320]
'0070c9e3    50                      push eax
'0070c9e4    8d55a4                  lea edx, dword ptr [ebp-5c]
'0070c9e7    52                      push edx
'0070c9e8    ffd3                    call ebx
Set var_83 = Nothing
'0070c9ea    50                      push eax

' *** Reference to "__vbaLateIdCallSt"
'0070c9eb    ff159c114000            call dword ptr [0040119c]
'0070c9f1    8d45c8                  lea eax, dword ptr [ebp-38]
'0070c9f4    50                      push eax
'0070c9f5    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070c9f8    51                      push ecx
'0070c9f9    8d55d0                  lea edx, dword ptr [ebp-30]
'0070c9fc    52                      push edx
'0070c9fd    6a03                    push 03

' *** Reference to "__vbaFreeStrList"
'0070c9ff    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , -4780, 0)
'0070ca05    83c45c                  add esp, 5c
'0070ca08    8d45a4                  lea eax, dword ptr [ebp-5c]
'0070ca0b    50                      push eax
'0070ca0c    8d4da8                  lea ecx, dword ptr [ebp-58]
'0070ca0f    51                      push ecx
'0070ca10    8d55ac                  lea edx, dword ptr [ebp-54]
'0070ca13    52                      push edx
'0070ca14    8d45b0                  lea eax, dword ptr [ebp-50]
'0070ca17    50                      push eax
'0070ca18    8d4db4                  lea ecx, dword ptr [ebp-4c]
'0070ca1b    51                      push ecx
'0070ca1c    8d55b8                  lea edx, dword ptr [ebp-48]
'0070ca1f    52                      push edx
'0070ca20    6a06                    push 06

' *** Reference to "__vbaFreeObjList"
'0070ca22    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_61, var_62, var_6, var_50, var_86, var_83)
'0070ca28    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'0070ca2e    50                      push eax
'0070ca2f    8d4d84                  lea ecx, dword ptr [ebp-7c]
'0070ca32    51                      push ecx
'0070ca33    8d5594                  lea edx, dword ptr [ebp-6c]
'0070ca36    52                      push edx
'0070ca37    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'0070ca39    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_121, var_119, var_91)
'0070ca3f    8b45d4                  mov eax, dword ptr [ebp-2c]
'0070ca42    8b08                    mov ecx, dword ptr [eax]
'0070ca44    83c42c                  add esp, 2c
'0070ca47    50                      push eax
'0070ca48    ff91c4000000            call dword ptr [ecx+000000c4]
'0070ca4e    dbe2                    fnclex
'0070ca50    85c0                    test ax, ax
'0070ca52    7d15                    jge 70ca69
'0070ca54    8b55d4                  mov edx, dword ptr [ebp-2c]
'0070ca57    68c4000000              push 000000c4
'0070ca5c    6830314300              push 00433130
'0070ca61    52                      push edx
'0070ca62    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070ca63    ff1580104000            call dword ptr [00401080]

' *** Reference to "__vbaLateIdCallLd"
'0070ca69    8b3d7c114000            mov edi, dword ptr [0040117c]

'ERROR: Two many next close:
End If
'0070ca6f    b801000000              mov eax, 00000001
'0070ca74    660345e8                add ax, word ptr [ebp-18]
var_num1 = 1 + var_41
'0070ca78    0f8018020000            jo 70cc96
'0070ca7e    8945e8                  mov dword ptr [ebp-18], eax
'0070ca81    e926f7ffff              jmp 70c1ac

'ERROR: Two many next close:
Loop
'0070ca86    b904000280              mov ecx, 80020004
'0070ca8b    b80a000000              mov eax, 0000000a
'0070ca90    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'0070ca96    898d7cffffff            mov dword ptr [ebp+ffffff7c], ecx
'0070ca9c    894d8c                  mov dword ptr [ebp-74], ecx
'0070ca9f    8d9514ffffff            lea edx, dword ptr [ebp+ffffff14]
'0070caa5    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0070caa8    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'0070caae    898574ffffff            mov dword ptr [ebp+ffffff74], eax
'0070cab4    894584                  mov dword ptr [ebp-7c], eax
'0070cab7    c7851cffffffd8644500    mov dword ptr [ebp+ffffff1c], 004564d8
'0070cac1    c78514ffffff08000000    mov dword ptr [ebp+ffffff14], 00000008

' *** Reference to "__vbaVarDup"
'0070cacb    ff1598124000            call dword ptr [00401298]
var_121 = ("Il faut au moins un personnage de sélectionné")
'0070cad1    8d8564ffffff            lea eax, dword ptr [ebp+ffffff64]
'0070cad7    50                      push eax
'0070cad8    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'0070cade    51                      push ecx
'0070cadf    8d5584                  lea edx, dword ptr [ebp-7c]
'0070cae2    52                      push edx
'0070cae3    6a00                    push 00
'0070cae5    8d4594                  lea eax, dword ptr [ebp-6c]
'0070cae8    50                      push eax

' *** Reference to "rtcMsgBox"
'0070cae9    ff15b8104000            call dword ptr [004010b8]
var_2344 = MsgBox(var_121, 0)
'0070caef    8d8d64ffffff            lea ecx, dword ptr [ebp+ffffff64]
'0070caf5    51                      push ecx
'0070caf6    8d9574ffffff            lea edx, dword ptr [ebp+ffffff74]
'0070cafc    52                      push edx
'0070cafd    8d4584                  lea eax, dword ptr [ebp-7c]
'0070cb00    50                      push eax
'0070cb01    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0070cb04    51                      push ecx
'0070cb05    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0070cb07    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_121, var_119, var_91, var_123)
'0070cb0d    83c414                  add esp, 14

'ERROR: Two many next close:
End If
'0070cb10    33ff                    xor edi, edi
var_num7 = Empty
'0070cb12    897dfc                  mov dword ptr [ebp-04], edi
'0070cb15    9b                      fwait
'0070cb16    6872cc7000              push 0070cc72
'0070cb1b    e948010000              jmp 70cc68

'ERROR: Two many next close:
End If
'0070cb20    b904000280              mov ecx, 80020004
'0070cb25    b80a000000              mov eax, 0000000a
'0070cb2a    898d6cffffff            mov dword ptr [ebp+ffffff6c], ecx
'0070cb30    898d7cffffff            mov dword ptr [ebp+ffffff7c], ecx
'0070cb36    894d8c                  mov dword ptr [ebp-74], ecx
'0070cb39    8d9514ffffff            lea edx, dword ptr [ebp+ffffff14]
'0070cb3f    8d4d94                  lea ecx, dword ptr [ebp-6c]
'0070cb42    898564ffffff            mov dword ptr [ebp+ffffff64], eax
'0070cb48    898574ffffff            mov dword ptr [ebp+ffffff74], eax
'0070cb4e    894584                  mov dword ptr [ebp-7c], eax
'0070cb51    c7851cffffff38654500    mov dword ptr [ebp+ffffff1c], 00456538
'0070cb5b    c78514ffffff08000000    mov dword ptr [ebp+ffffff14], 00000008

' *** Reference to "__vbaVarDup"
'0070cb65    ff1598124000            call dword ptr [00401298]
var_121 = ("Le FP doit être un réél strictement supérieur à 0")
'0070cb6b    8d9564ffffff            lea edx, dword ptr [ebp+ffffff64]
'0070cb71    52                      push edx
'0070cb72    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'0070cb78    50                      push eax
'0070cb79    8d4d84                  lea ecx, dword ptr [ebp-7c]
'0070cb7c    51                      push ecx
'0070cb7d    57                      push edi
'0070cb7e    8d5594                  lea edx, dword ptr [ebp-6c]
'0070cb81    52                      push edx

' *** Reference to "rtcMsgBox"
'0070cb82    ff15b8104000            call dword ptr [004010b8]
var_2114 = MsgBox(var_121, 0)
'0070cb88    8d8564ffffff            lea eax, dword ptr [ebp+ffffff64]
'0070cb8e    50                      push eax
'0070cb8f    8d8d74ffffff            lea ecx, dword ptr [ebp+ffffff74]
'0070cb95    51                      push ecx
'0070cb96    8d5584                  lea edx, dword ptr [ebp-7c]
'0070cb99    52                      push edx
'0070cb9a    8d4594                  lea eax, dword ptr [ebp-6c]
'0070cb9d    50                      push eax
'0070cb9e    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0070cba0    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_121, var_119, var_91, var_123)
'0070cba6    8b0e                    mov ecx, dword ptr [esi]
'0070cba8    83c414                  add esp, 14
'0070cbab    56                      push esi
'0070cbac    ff9104030000            call dword ptr [ecx+00000304]
'0070cbb2    50                      push eax
'0070cbb3    8d55b8                  lea edx, dword ptr [ebp-48]
'0070cbb6    52                      push edx
'0070cbb7    ffd3                    call ebx
Set var_61 = 
'0070cbb9    8bf0                    mov esi, eax
'0070cbbb    8b06                    mov eax, dword ptr [esi]
'0070cbbd    56                      push esi
'0070cbbe    ff9004020000            call dword ptr [eax+00000204]
'0070cbc4    dbe2                    fnclex
'0070cbc6    3bc7                    cmp eax, edi
'0070cbc8    7d12                    jge 70cbdc

If (var_121 < 0) Then
'0070cbca    6804020000              push 00000204
'0070cbcf    68d00d4300              push 00430dd0
'0070cbd4    56                      push esi
'0070cbd5    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070cbd6    ff1580104000            call dword ptr [00401080]
    
End If
'0070cbdc    8d4db8                  lea ecx, dword ptr [ebp-48]

' *** Reference to "__vbaFreeObj"
'0070cbdf    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_61)
'0070cbe5    e928ffffff              jmp 70cb12
'0070cbea    8d4dbc                  lea ecx, dword ptr [ebp-44]
'0070cbed    51                      push ecx
'0070cbee    8d55c0                  lea edx, dword ptr [ebp-40]
'0070cbf1    52                      push edx
'0070cbf2    8d45c4                  lea eax, dword ptr [ebp-3c]
'0070cbf5    50                      push eax
'0070cbf6    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0070cbf9    51                      push ecx
'0070cbfa    8d55cc                  lea edx, dword ptr [ebp-34]
'0070cbfd    52                      push edx
'0070cbfe    8d45d0                  lea eax, dword ptr [ebp-30]
'0070cc01    50                      push eax
'0070cc02    6a06                    push 06

' *** Reference to "__vbaFreeStrList"
'0070cc04    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , -4780, 0, -4784, -4788, -4792)
'0070cc0a    8d4da4                  lea ecx, dword ptr [ebp-5c]
'0070cc0d    51                      push ecx
'0070cc0e    8d55a8                  lea edx, dword ptr [ebp-58]
'0070cc11    52                      push edx
'0070cc12    8d45ac                  lea eax, dword ptr [ebp-54]
'0070cc15    50                      push eax
'0070cc16    8d4db0                  lea ecx, dword ptr [ebp-50]
'0070cc19    51                      push ecx
'0070cc1a    8d55b4                  lea edx, dword ptr [ebp-4c]
'0070cc1d    52                      push edx
'0070cc1e    8d45b8                  lea eax, dword ptr [ebp-48]
'0070cc21    50                      push eax
'0070cc22    6a06                    push 06

' *** Reference to "__vbaFreeObjList"
'0070cc24    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_61, var_62, var_6, var_50, var_86, var_83)
'0070cc2a    8d8d24ffffff            lea ecx, dword ptr [ebp+ffffff24]
'0070cc30    51                      push ecx
'0070cc31    8d9534ffffff            lea edx, dword ptr [ebp+ffffff34]
'0070cc37    52                      push edx
'0070cc38    8d8544ffffff            lea eax, dword ptr [ebp+ffffff44]
'0070cc3e    50                      push eax
'0070cc3f    8d8d54ffffff            lea ecx, dword ptr [ebp+ffffff54]
'0070cc45    51                      push ecx
'0070cc46    8d9564ffffff            lea edx, dword ptr [ebp+ffffff64]
'0070cc4c    52                      push edx
'0070cc4d    8d8574ffffff            lea eax, dword ptr [ebp+ffffff74]
'0070cc53    50                      push eax
'0070cc54    8d4d84                  lea ecx, dword ptr [ebp-7c]
'0070cc57    51                      push ecx
'0070cc58    8d5594                  lea edx, dword ptr [ebp-6c]
'0070cc5b    52                      push edx
'0070cc5c    6a08                    push 08

' *** Reference to "__vbaFreeVarList"
'0070cc5e    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_121, var_119, var_91, var_123, var_124, var_125, var_317, var_432)
'0070cc64    83c45c                  add esp, 5c
'0070cc67    c3                      ret
'0070cc68    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaFreeObj"
'0070cc6b    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_44)
'0070cc71    c3                      ret
'0070cc72    8b4508                  mov eax, dword ptr [ebp+08]
'0070cc75    8b08                    mov ecx, dword ptr [eax]
'0070cc77    50                      push eax
'0070cc78    ff5108                  call dword ptr [ecx+08]
'0070cc7b    8b45fc                  mov eax, dword ptr [ebp-04]
'0070cc7e    8b4dec                  mov ecx, dword ptr [ebp-14]
'0070cc81    5f                      pop edi
'0070cc82    5e                      pop esi
    'Reference to '__except_list'
'0070cc83    64890d00000000          mov dword ptr fs:[00000000], ecx
'0070cc8a    5b                      pop ebx
'0070cc8b    8be5                    mov esp, ebp
'0070cc8d    5d                      pop ebp
'0070cc8e    c20400                  ret 0004


End Sub


'Event for BtnFermer
Private Sub BtnFermer_Click()
'0070e310    55                      push ebp
'0070e311    8bec                    mov ebp, esp
'0070e313    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0070e316    6866724000              push 00407266
'0070e31b    64a100000000            mov ax, word ptr fs:[00000000]
'0070e321    50                      push eax
    'Reference to '__except_list'
'0070e322    64892500000000          mov dword ptr fs:[00000000], esp
'0070e329    83ec18                  sub esp, 18
'0070e32c    53                      push ebx
'0070e32d    56                      push esi
'0070e32e    57                      push edi
'0070e32f    8965f4                  mov dword ptr [ebp-0c], esp
'0070e332    c745f8506c4000          mov dword ptr [ebp-08], 00406c50
'0070e339    8b7d08                  mov edi, dword ptr [ebp+08]
'0070e33c    8bc7                    mov eax, edi
'0070e33e    83e001                  and eax, 01
'0070e341    8945fc                  mov dword ptr [ebp-04], eax
'0070e344    83e7fe                  and edi, fffffffe
'0070e347    8b0f                    mov ecx, dword ptr [edi]
'0070e349    57                      push edi
'0070e34a    897d08                  mov dword ptr [ebp+08], edi
'0070e34d    ff5104                  call dword ptr [ecx+04]
'0070e350    a124be7200              mov ax, word ptr [0072be24]
'0070e355    33db                    xor ebx, ebx
'0070e357    3bc3                    cmp eax, ebx
'0070e359    895de8                  mov dword ptr [ebp-18], ebx
'0070e35c    7510                    jne 70e36e

If (0 = 0) Then
'0070e35e    6824be7200              push 0072be24
'0070e363    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'0070e368    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
End If
'0070e36e    8b3524be7200            mov esi, dword ptr [0072be24]
'0070e374    8b16                    mov edx, dword ptr [esi]
'0070e376    57                      push edi
'0070e377    8d45e8                  lea eax, dword ptr [ebp-18]
'0070e37a    50                      push eax
'0070e37b    8955d4                  mov dword ptr [ebp-2c], edx

' *** Reference to "__vbaObjSetAddref"
'0070e37e    ff15c8104000            call dword ptr [004010c8]
Set var_41 = Me
'0070e384    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'0070e387    50                      push eax
'0070e388    56                      push esi
'0070e389    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_41)
'0070e38c    dbe2                    fnclex
'0070e38e    3bc3                    cmp eax, ebx
'0070e390    7d0f                    jge 70e3a1
'0070e392    6a10                    push 10
'0070e394    6860004300              push 00430060
'0070e399    56                      push esi
'0070e39a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070e39b    ff1580104000            call dword ptr [00401080]
'0070e3a1    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0070e3a4    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0070e3aa    895dfc                  mov dword ptr [ebp-04], ebx
'0070e3ad    68bfe37000              push 0070e3bf
'0070e3b2    eb0a                    jmp 70e3be
'0070e3b4    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'0070e3b7    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'0070e3bd    c3                      ret
'0070e3be    c3                      ret
'0070e3bf    8b4508                  mov eax, dword ptr [ebp+08]
'0070e3c2    8b10                    mov edx, dword ptr [eax]
'0070e3c4    50                      push eax
'0070e3c5    ff5208                  call dword ptr [edx+08]
'0070e3c8    8b45fc                  mov eax, dword ptr [ebp-04]
'0070e3cb    8b4dec                  mov ecx, dword ptr [ebp-14]
'0070e3ce    5f                      pop edi
'0070e3cf    5e                      pop esi
    'Reference to '__except_list'
'0070e3d0    64890d00000000          mov dword ptr fs:[00000000], ecx
'0070e3d7    5b                      pop ebx
'0070e3d8    8be5                    mov esp, ebp
'0070e3da    5d                      pop ebp
'0070e3db    c20400                  ret 0004


End Sub


'Event for vspersonnage
Private Sub vspersonnage_Event9()
'0070f5c0    55                      push ebp
'0070f5c1    8bec                    mov ebp, esp
'0070f5c3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0070f5c6    6866724000              push 00407266
'0070f5cb    64a100000000            mov ax, word ptr fs:[00000000]
'0070f5d1    50                      push eax
    'Reference to '__except_list'
'0070f5d2    64892500000000          mov dword ptr fs:[00000000], esp
'0070f5d9    83ec20                  sub esp, 20
'0070f5dc    53                      push ebx
'0070f5dd    56                      push esi
'0070f5de    57                      push edi
'0070f5df    8965f4                  mov dword ptr [ebp-0c], esp
'0070f5e2    c745f8e06c4000          mov dword ptr [ebp-08], 00406ce0
'0070f5e9    8b7508                  mov esi, dword ptr [ebp+08]
'0070f5ec    8bc6                    mov eax, esi
'0070f5ee    83e001                  and eax, 01
'0070f5f1    8945fc                  mov dword ptr [ebp-04], eax
'0070f5f4    83e6fe                  and esi, fffffffe
'0070f5f7    8b0e                    mov ecx, dword ptr [esi]
'0070f5f9    56                      push esi
'0070f5fa    897508                  mov dword ptr [ebp+08], esi
'0070f5fd    ff5104                  call dword ptr [ecx+04]
'0070f600    8b16                    mov edx, dword ptr [esi]
'0070f602    33ff                    xor edi, edi
'0070f604    57                      push edi
'0070f605    6a11                    push 11
'0070f607    56                      push esi
'0070f608    897de4                  mov dword ptr [ebp-1c], edi
'0070f60b    897dd4                  mov dword ptr [ebp-2c], edi
'0070f60e    ff9220030000            call dword ptr [edx+00000320]
'0070f614    50                      push eax
'0070f615    8d45e4                  lea eax, dword ptr [ebp-1c]
'0070f618    50                      push eax

' *** Reference to "__vbaObjSet"
'0070f619    ff15b4104000            call dword ptr [004010b4]
Set var_40 = Me
'0070f61f    50                      push eax
'0070f620    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0070f623    51                      push ecx

' *** Reference to "__vbaLateIdCallLd"
'0070f624    ff157c114000            call dword ptr [0040117c]
var_44 = var_40.UNK_frmExperience_17
'0070f62a    83c410                  add esp, 10
'0070f62d    50                      push eax

' *** Reference to "__vbaI4Var"
'0070f62e    ff157c124000            call dword ptr [0040127c]
'0070f634    8bc8                    mov ecx, eax

' *** Reference to "__vbaI2I4"
'0070f636    ff1550114000            call dword ptr [00401150]
'0070f63c    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'0070f63f    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)
'0070f645    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaFreeVar"
'0070f648    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_44)
'0070f64e    897dfc                  mov dword ptr [ebp-04], edi
'0070f651    686cf67000              push 0070f66c
'0070f656    eb13                    jmp 70f66b
'0070f658    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'0070f65b    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)
'0070f661    8d4dd4                  lea ecx, dword ptr [ebp-2c]

' *** Reference to "__vbaFreeVar"
'0070f664    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_44)
'0070f66a    c3                      ret
'0070f66b    c3                      ret
'0070f66c    8b4508                  mov eax, dword ptr [ebp+08]
'0070f66f    8b10                    mov edx, dword ptr [eax]
'0070f671    50                      push eax
'0070f672    ff5208                  call dword ptr [edx+08]
'0070f675    8b45fc                  mov eax, dword ptr [ebp-04]
'0070f678    8b4dec                  mov ecx, dword ptr [ebp-14]
'0070f67b    5f                      pop edi
'0070f67c    5e                      pop esi
    'Reference to '__except_list'
'0070f67d    64890d00000000          mov dword ptr fs:[00000000], ecx
'0070f684    5b                      pop ebx
'0070f685    8be5                    mov esp, ebp
'0070f687    5d                      pop ebp
'0070f688    c20400                  ret 0004


End Sub


Private Sub vspersonnage_Event37()
'0070f030    55                      push ebp
'0070f031    8bec                    mov ebp, esp
'0070f033    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'0070f036    6866724000              push 00407266
'0070f03b    64a100000000            mov ax, word ptr fs:[00000000]
'0070f041    50                      push eax
    'Reference to '__except_list'
'0070f042    64892500000000          mov dword ptr fs:[00000000], esp
'0070f049    81ec2c010000            sub esp, 0000012c
'0070f04f    53                      push ebx
'0070f050    56                      push esi
'0070f051    57                      push edi
'0070f052    8965ec                  mov dword ptr [ebp-14], esp
'0070f055    c745f0b86c4000          mov dword ptr [ebp-10], 00406cb8
'0070f05c    8b7d08                  mov edi, dword ptr [ebp+08]
'0070f05f    8bc7                    mov eax, edi
'0070f061    83e001                  and eax, 01
'0070f064    8945f4                  mov dword ptr [ebp-0c], eax
'0070f067    83e7fe                  and edi, fffffffe
'0070f06a    897d08                  mov dword ptr [ebp+08], edi
'0070f06d    33f6                    xor esi, esi
'0070f06f    8975f8                  mov dword ptr [ebp-08], esi
'0070f072    8b0f                    mov ecx, dword ptr [edi]
'0070f074    57                      push edi
'0070f075    ff5104                  call dword ptr [ecx+04]
'0070f078    8975e0                  mov dword ptr [ebp-20], esi
'0070f07b    8975dc                  mov dword ptr [ebp-24], esi
'0070f07e    8975d8                  mov dword ptr [ebp-28], esi
'0070f081    8975d4                  mov dword ptr [ebp-2c], esi
'0070f084    8975d0                  mov dword ptr [ebp-30], esi
'0070f087    8975cc                  mov dword ptr [ebp-34], esi
'0070f08a    8975c8                  mov dword ptr [ebp-38], esi
'0070f08d    8975c4                  mov dword ptr [ebp-3c], esi
'0070f090    8975c0                  mov dword ptr [ebp-40], esi
'0070f093    8975b0                  mov dword ptr [ebp-50], esi
'0070f096    8975a0                  mov dword ptr [ebp-60], esi
'0070f099    897590                  mov dword ptr [ebp-70], esi
'0070f09c    897580                  mov dword ptr [ebp-80], esi
'0070f09f    89b570ffffff            mov dword ptr [ebp+ffffff70], esi
'0070f0a5    89b560ffffff            mov dword ptr [ebp+ffffff60], esi
'0070f0ab    89b550ffffff            mov dword ptr [ebp+ffffff50], esi
'0070f0b1    89b540ffffff            mov dword ptr [ebp+ffffff40], esi
'0070f0b7    89b530ffffff            mov dword ptr [ebp+ffffff30], esi
'0070f0bd    89b520ffffff            mov dword ptr [ebp+ffffff20], esi
'0070f0c3    89b510ffffff            mov dword ptr [ebp+ffffff10], esi
'0070f0c9    89b500ffffff            mov dword ptr [ebp+ffffff00], esi
'0070f0cf    89b5f0feffff            mov dword ptr [ebp+fffffef0], esi
'0070f0d5    89b5e0feffff            mov dword ptr [ebp+fffffee0], esi
'0070f0db    89b5dcfeffff            mov dword ptr [ebp+fffffedc], esi
'0070f0e1    66393528b07200          cmp word ptr [0072b028], si
'0070f0e8    7508                    jne 70f0f2
'0070f0ea    6a01                    push 01

' *** Reference to "__vbaOnError"
'0070f0ec    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'0070f0f2    39750c                  cmp dword ptr [ebp+0c], esi
'0070f0f5    8b4510                  mov eax, dword ptr [ebp+10]
'0070f0f8    0f8ea2030000            jle 70f4a0

If (arg_0 > 0) Then
'0070f0fe    83f805                  cmp eax, 05
'0070f101    740e                    je 70f111
    
    If (    arg_1 <> 5) Then
'0070f103    83f803                  cmp eax, 03
'0070f106    7409                    je 70f111
    
    If (    arg_1 <> 3) Then
'0070f108    83f806                  cmp eax, 06
'0070f10b    0f858f030000            jne 70f4a0
    
    If (    arg_1 = 6) Then
'0070f111    b8cc134300              mov eax, 004313cc
'0070f116    898528ffffff            mov dword ptr [ebp+ffffff28], eax
'0070f11c    b908000000              mov ecx, 00000008
'0070f121    898d20ffffff            mov dword ptr [ebp+ffffff20], ecx
'0070f127    83ec10                  sub esp, 10
'0070f12a    8bd4                    mov edx, esp
'0070f12c    890a                    mov dword ptr [edx], ecx
'0070f12e    8b8d24ffffff            mov ecx, dword ptr [ebp+ffffff24]
'0070f134    894a04                  mov dword ptr [edx+04], ecx
'0070f137    894208                  mov dword ptr [edx+08], eax
'0070f13a    8b852cffffff            mov eax, dword ptr [ebp+ffffff2c]
'0070f140    89420c                  mov dword ptr [edx+0c], eax
'0070f143    6a48                    push 48
'0070f145    8b0f                    mov ecx, dword ptr [edi]
'0070f147    57                      push edi
'0070f148    ff9120030000            call dword ptr [ecx+00000320]
'0070f14e    50                      push eax
'0070f14f    8d55c4                  lea edx, dword ptr [ebp-3c]
'0070f152    52                      push edx

' *** Reference to "__vbaObjSet"
'0070f153    ff15b4104000            call dword ptr [004010b4]
    Set var_9 = Nothing
'0070f159    50                      push eax

' *** Reference to "__vbaLateIdSt"
'0070f15a    ff15ec124000            call dword ptr [004012ec]
    var_9.[UNMANAGED] = vbNullChar
'0070f160    8d4dc4                  lea ecx, dword ptr [ebp-3c]

' *** Reference to "__vbaFreeObj"
'0070f163    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_9)
'0070f169    e9a7030000              jmp 70f515

' *** Reference to "rtcErrObj"
'0070f16e    8b1d6c124000            mov ebx, dword ptr [0040126c]
'0070f174    ffd3                    call ebx
'0070f176    50                      push eax
'0070f177    8d55c4                  lea edx, dword ptr [ebp-3c]
'0070f17a    52                      push edx

' *** Reference to "__vbaObjSet"
'0070f17b    8b3db4104000            mov edi, dword ptr [004010b4]
'0070f181    ffd7                    call edi
    Set var_9 = Err
'0070f183    8bf0                    mov esi, eax
'0070f185    8b06                    mov eax, dword ptr [esi]
'0070f187    8d4de0                  lea ecx, dword ptr [ebp-20]
'0070f18a    51                      push ecx
'0070f18b    56                      push esi
'0070f18c    ff502c                  call dword ptr [eax+2c]
    var_3 = var_9.Description
'0070f18f    dbe2                    fnclex
'0070f191    85c0                    test ax, ax
'0070f193    7d0f                    jge 70f1a4
'0070f195    6a2c                    push 2c
'0070f197    685c084300              push 0043085c
'0070f19c    56                      push esi
'0070f19d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070f19e    ff1580104000            call dword ptr [00401080]
'0070f1a4    ffd3                    call ebx
'0070f1a6    50                      push eax
'0070f1a7    8d55c0                  lea edx, dword ptr [ebp-40]
'0070f1aa    52                      push edx
'0070f1ab    ffd7                    call edi
    Set var_5 = Err
'0070f1ad    8bf0                    mov esi, eax
'0070f1af    8b06                    mov eax, dword ptr [esi]
'0070f1b1    8d8ddcfeffff            lea ecx, dword ptr [ebp+fffffedc]
'0070f1b7    51                      push ecx
'0070f1b8    56                      push esi
'0070f1b9    ff501c                  call dword ptr [eax+1c]
    var_10 = var_5.Number
'0070f1bc    dbe2                    fnclex
'0070f1be    85c0                    test ax, ax
'0070f1c0    7d0f                    jge 70f1d1
'0070f1c2    6a1c                    push 1c
'0070f1c4    685c084300              push 0043085c
'0070f1c9    56                      push esi
'0070f1ca    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070f1cb    ff1580104000            call dword ptr [00401080]
'0070f1d1    b904000280              mov ecx, 80020004
'0070f1d6    894d98                  mov dword ptr [ebp-68], ecx
'0070f1d9    b80a000000              mov eax, 0000000a
'0070f1de    894590                  mov dword ptr [ebp-70], eax
'0070f1e1    894da8                  mov dword ptr [ebp-58], ecx
'0070f1e4    8945a0                  mov dword ptr [ebp-60], eax
'0070f1e7    c78528ffffff24b07200    mov dword ptr [ebp+ffffff28], 0072b024
'0070f1f1    c78520ffffff08400000    mov dword ptr [ebp+ffffff20], 00004008
'0070f1fb    6814084300              push 00430814
'0070f200    8b55e0                  mov edx, dword ptr [ebp-20]
'0070f203    52                      push edx

' *** Reference to "__vbaStrCat"
'0070f204    8b3570104000            mov esi, dword ptr [00401070]
'0070f20a    ffd6                    call esi
    var_11 = ("L'erreur suivante s'est produite : ") & (var_3)
'0070f20c    8bd0                    mov edx, eax
'0070f20e    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'0070f211    8b3dd0124000            mov edi, dword ptr [004012d0]
'0070f217    ffd7                    call edi
'0070f219    50                      push eax
'0070f21a    6870084300              push 00430870
'0070f21f    ffd6                    call esi
    var_127 = (var_11) & (vbCrLf)
'0070f221    8bd0                    mov edx, eax
'0070f223    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0070f226    ffd7                    call edi
'0070f228    50                      push eax
'0070f229    6870084300              push 00430870
'0070f22e    ffd6                    call esi
    var_15 = (var_127) & (vbCrLf)
'0070f230    8bd0                    mov edx, eax
'0070f232    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0070f235    ffd7                    call edi
'0070f237    50                      push eax
'0070f238    6880084300              push 00430880
'0070f23d    ffd6                    call esi
    var_16 = (var_15) & (" numéro d'erreur (")
'0070f23f    8bd0                    mov edx, eax
'0070f241    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0070f244    ffd7                    call edi
'0070f246    50                      push eax
'0070f247    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'0070f24d    50                      push eax

' *** Reference to "__vbaStrI4"
'0070f24e    ff1520104000            call dword ptr [00401020]
'0070f254    8bd0                    mov edx, eax
'0070f256    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070f259    ffd7                    call edi
'0070f25b    50                      push eax
'0070f25c    ffd6                    call esi
    var_17 = (var_16) & (CStr(var_10))
'0070f25e    8bd0                    mov edx, eax
'0070f260    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0070f263    ffd7                    call edi
'0070f265    50                      push eax
'0070f266    68ac084300              push 004308ac
'0070f26b    ffd6                    call esi
    var_129 = (var_17) & (" )")
'0070f26d    8945b8                  mov dword ptr [ebp-48], eax
'0070f270    bf08000000              mov edi, 00000008
'0070f275    897db0                  mov dword ptr [ebp-50], edi
'0070f278    8d4d90                  lea ecx, dword ptr [ebp-70]
'0070f27b    51                      push ecx
'0070f27c    8d55a0                  lea edx, dword ptr [ebp-60]
'0070f27f    52                      push edx
'0070f280    8d8520ffffff            lea eax, dword ptr [ebp+ffffff20]
'0070f286    50                      push eax
'0070f287    6a10                    push 10
'0070f289    8d4db0                  lea ecx, dword ptr [ebp-50]
'0070f28c    51                      push ecx

' *** Reference to "rtcMsgBox"
'0070f28d    ff15b8104000            call dword ptr [004010b8]
    var_285 = MsgBox(var_129, 16, vbNullString)
'0070f293    8d55c8                  lea edx, dword ptr [ebp-38]
'0070f296    52                      push edx
'0070f297    8d45cc                  lea eax, dword ptr [ebp-34]
'0070f29a    50                      push eax
'0070f29b    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0070f29e    51                      push ecx
'0070f29f    8d55d4                  lea edx, dword ptr [ebp-2c]
'0070f2a2    52                      push edx
'0070f2a3    8d45d8                  lea eax, dword ptr [ebp-28]
'0070f2a6    50                      push eax
'0070f2a7    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0070f2aa    51                      push ecx
'0070f2ab    8d55e0                  lea edx, dword ptr [ebp-20]
'0070f2ae    52                      push edx
'0070f2af    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0070f2b1    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( 0, -4524, -4528, -4532, -4536, -4540, -4544)
'0070f2b7    8d45c0                  lea eax, dword ptr [ebp-40]
'0070f2ba    50                      push eax
'0070f2bb    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0070f2be    51                      push ecx
'0070f2bf    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070f2c1    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_9, var_5)
'0070f2c7    8d5590                  lea edx, dword ptr [ebp-70]
'0070f2ca    52                      push edx
'0070f2cb    8d45a0                  lea eax, dword ptr [ebp-60]
'0070f2ce    50                      push eax
'0070f2cf    8d4db0                  lea ecx, dword ptr [ebp-50]
'0070f2d2    51                      push ecx
'0070f2d3    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'0070f2d5    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7, var_8)
'0070f2db    83c43c                  add esp, 3c
'0070f2de    8d55b0                  lea edx, dword ptr [ebp-50]
'0070f2e1    52                      push edx

' *** Reference to "rtcGetPresentDate"
'0070f2e2    ff15f4124000            call dword ptr [004012f4]
    var_129 = Now()
'0070f2e8    c78528ffffffb8084300    mov dword ptr [ebp+ffffff28], 004308b8
'0070f2f2    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'0070f2f8    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'0070f2fe    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaVarDup"
'0070f301    ff1598124000            call dword ptr [00401298]
    var_7 = ("dd/MM/yyyy hh:mm:ss")
'0070f307    6a01                    push 01
'0070f309    6a01                    push 01
'0070f30b    8d45a0                  lea eax, dword ptr [ebp-60]
'0070f30e    50                      push eax
'0070f30f    8d4db0                  lea ecx, dword ptr [ebp-50]
'0070f312    51                      push ecx
'0070f313    8d5590                  lea edx, dword ptr [ebp-70]
'0070f316    52                      push edx

' *** Reference to "rtcVarFromFormatVar"
'0070f317    ff1574104000            call dword ptr [00401074]
'0070f31d    ffd3                    call ebx
'0070f31f    50                      push eax
'0070f320    8d45c4                  lea eax, dword ptr [ebp-3c]
'0070f323    50                      push eax

' *** Reference to "__vbaObjSet"
'0070f324    ff15b4104000            call dword ptr [004010b4]
    Set var_9 = Err
'0070f32a    8bf0                    mov esi, eax
'0070f32c    8b0e                    mov ecx, dword ptr [esi]
'0070f32e    8d95dcfeffff            lea edx, dword ptr [ebp+fffffedc]
'0070f334    52                      push edx
'0070f335    56                      push esi
'0070f336    ff511c                  call dword ptr [ecx+1c]
    var_10 = var_9.Number
'0070f339    dbe2                    fnclex
'0070f33b    85c0                    test ax, ax
'0070f33d    7d0f                    jge 70f34e
'0070f33f    6a1c                    push 1c
'0070f341    685c084300              push 0043085c
'0070f346    56                      push esi
'0070f347    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070f348    ff1580104000            call dword ptr [00401080]
'0070f34e    ffd3                    call ebx
'0070f350    50                      push eax
'0070f351    8d45c0                  lea eax, dword ptr [ebp-40]
'0070f354    50                      push eax

' *** Reference to "__vbaObjSet"
'0070f355    ff15b4104000            call dword ptr [004010b4]
    Set var_5 = Err
'0070f35b    8bf0                    mov esi, eax
'0070f35d    8b0e                    mov ecx, dword ptr [esi]
'0070f35f    8d55e0                  lea edx, dword ptr [ebp-20]
'0070f362    52                      push edx
'0070f363    56                      push esi
'0070f364    ff512c                  call dword ptr [ecx+2c]
    var_3 = var_5.Description
'0070f367    dbe2                    fnclex
'0070f369    85c0                    test ax, ax
'0070f36b    7d0f                    jge 70f37c
'0070f36d    6a2c                    push 2c
'0070f36f    685c084300              push 0043085c
'0070f374    56                      push esi
'0070f375    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070f376    ff1580104000            call dword ptr [00401080]
'0070f37c    c78518ffffffe4084300    mov dword ptr [ebp+ffffff18], 004308e4
'0070f386    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'0070f38c    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'0070f392    898508ffffff            mov dword ptr [ebp+ffffff08], eax
'0070f398    c78500ffffff03000000    mov dword ptr [ebp+ffffff00], 00000003
'0070f3a2    c785f8feffff08094300    mov dword ptr [ebp+fffffef8], 00430908
'0070f3ac    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'0070f3b2    8b45e0                  mov eax, dword ptr [ebp-20]
'0070f3b5    c745e000000000          mov dword ptr [ebp-20], 00000000
'0070f3bc    898558ffffff            mov dword ptr [ebp+ffffff58], eax
'0070f3c2    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'0070f3c8    c785e8fefffff8664500    mov dword ptr [ebp+fffffee8], 004566f8
'0070f3d2    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'0070f3d8    8d4d90                  lea ecx, dword ptr [ebp-70]
'0070f3db    51                      push ecx
'0070f3dc    8d9510ffffff            lea edx, dword ptr [ebp+ffffff10]
'0070f3e2    52                      push edx
'0070f3e3    8d4580                  lea eax, dword ptr [ebp-80]
'0070f3e6    50                      push eax

' *** Reference to "__vbaVarCat"
'0070f3e7    8b3508124000            mov esi, dword ptr [00401208]
'0070f3ed    ffd6                    call esi
'0070f3ef    50                      push eax
'0070f3f0    8d8d00ffffff            lea ecx, dword ptr [ebp+ffffff00]
'0070f3f6    51                      push ecx
'0070f3f7    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0070f3fd    52                      push edx
'0070f3fe    ffd6                    call esi
'0070f400    50                      push eax
'0070f401    8d85f0feffff            lea eax, dword ptr [ebp+fffffef0]
'0070f407    50                      push eax
'0070f408    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'0070f40e    51                      push ecx
'0070f40f    ffd6                    call esi
'0070f411    50                      push eax
'0070f412    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'0070f418    52                      push edx
'0070f419    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'0070f41f    50                      push eax
'0070f420    ffd6                    call esi
'0070f422    50                      push eax
'0070f423    8d8de0feffff            lea ecx, dword ptr [ebp+fffffee0]
'0070f429    51                      push ecx
'0070f42a    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'0070f430    52                      push edx
'0070f431    ffd6                    call esi
'0070f433    50                      push eax
'0070f434    33c0                    xor eax, eax
'0070f436    66a12ab07200            mov eax, dword ptr [0072b02a]
'0070f43c    50                      push eax
'0070f43d    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'0070f442    ff15b8114000            call dword ptr [004011b8]
    Print #0, 
'0070f448    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0070f44b    51                      push ecx
'0070f44c    8d55c4                  lea edx, dword ptr [ebp-3c]
'0070f44f    52                      push edx
'0070f450    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070f452    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_9, var_5)
'0070f458    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'0070f45e    50                      push eax
'0070f45f    8d8d40ffffff            lea ecx, dword ptr [ebp+ffffff40]
'0070f465    51                      push ecx
'0070f466    8d9550ffffff            lea edx, dword ptr [ebp+ffffff50]
'0070f46c    52                      push edx
'0070f46d    8d8560ffffff            lea eax, dword ptr [ebp+ffffff60]
'0070f473    50                      push eax
'0070f474    8d8d70ffffff            lea ecx, dword ptr [ebp+ffffff70]
'0070f47a    51                      push ecx
'0070f47b    8d5580                  lea edx, dword ptr [ebp-80]
'0070f47e    52                      push edx
'0070f47f    8d4590                  lea eax, dword ptr [ebp-70]
'0070f482    50                      push eax
'0070f483    8d4da0                  lea ecx, dword ptr [ebp-60]
'0070f486    51                      push ecx
'0070f487    8d55b0                  lea edx, dword ptr [ebp-50]
'0070f48a    52                      push edx
'0070f48b    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0070f48d    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'0070f493    83c440                  add esp, 40
'0070f496    6a00                    push 00

' *** Reference to "__vbaResume"
'0070f498    ff1568104000            call dword ptr [00401068]
'0070f49e    eb75                    jmp 70f515
    Resume handler_70F515
End If
'0070f4a0    3bc6                    cmp eax, esi
'0070f4a2    7471                    je 70f515

If (arg_1 <> 0) Then
'0070f4a4    b904000280              mov ecx, 80020004
'0070f4a9    894d88                  mov dword ptr [ebp-78], ecx
'0070f4ac    b80a000000              mov eax, 0000000a
'0070f4b1    894580                  mov dword ptr [ebp-80], eax
'0070f4b4    894d98                  mov dword ptr [ebp-68], ecx
'0070f4b7    894590                  mov dword ptr [ebp-70], eax
'0070f4ba    894da8                  mov dword ptr [ebp-58], ecx
'0070f4bd    8945a0                  mov dword ptr [ebp-60], eax
'0070f4c0    c78528ffffffa0664500    mov dword ptr [ebp+ffffff28], 004566a0
'0070f4ca    c78520ffffff08000000    mov dword ptr [ebp+ffffff20], 00000008
'0070f4d4    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'0070f4da    8d4db0                  lea ecx, dword ptr [ebp-50]

' *** Reference to "__vbaVarDup"
'0070f4dd    ff1598124000            call dword ptr [00401298]
    var_6 = ("Vous ne pouvez pas changer cette céllule")
'0070f4e3    8d4580                  lea eax, dword ptr [ebp-80]
'0070f4e6    50                      push eax
'0070f4e7    8d4d90                  lea ecx, dword ptr [ebp-70]
'0070f4ea    51                      push ecx
'0070f4eb    8d55a0                  lea edx, dword ptr [ebp-60]
'0070f4ee    52                      push edx
'0070f4ef    56                      push esi
'0070f4f0    8d45b0                  lea eax, dword ptr [ebp-50]
'0070f4f3    50                      push eax

' *** Reference to "rtcMsgBox"
'0070f4f4    ff15b8104000            call dword ptr [004010b8]
    var_873 = MsgBox(var_6, 0)
'0070f4fa    8d4d80                  lea ecx, dword ptr [ebp-80]
'0070f4fd    51                      push ecx
'0070f4fe    8d5590                  lea edx, dword ptr [ebp-70]
'0070f501    52                      push edx
'0070f502    8d45a0                  lea eax, dword ptr [ebp-60]
'0070f505    50                      push eax
'0070f506    8d4db0                  lea ecx, dword ptr [ebp-50]
'0070f509    51                      push ecx
'0070f50a    6a04                    push 04

' *** Reference to "__vbaFreeVarList"
'0070f50c    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_7, var_8, var_18)
'0070f512    83c414                  add esp, 14
    
End If

' *** Reference to "__vbaExitProc"
'0070f515    ff15a0104000            call dword ptr [004010a0]
'0070f51b    6896f57000              push 0070f596
'0070f520    eb73                    jmp 70f595
'0070f522    8d55c8                  lea edx, dword ptr [ebp-38]
'0070f525    52                      push edx
'0070f526    8d45cc                  lea eax, dword ptr [ebp-34]
'0070f529    50                      push eax
'0070f52a    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0070f52d    51                      push ecx
'0070f52e    8d55d4                  lea edx, dword ptr [ebp-2c]
'0070f531    52                      push edx
'0070f532    8d45d8                  lea eax, dword ptr [ebp-28]
'0070f535    50                      push eax
'0070f536    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0070f539    51                      push ecx
'0070f53a    8d55e0                  lea edx, dword ptr [ebp-20]
'0070f53d    52                      push edx
'0070f53e    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0070f540    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_3, -4524, -4528, -4532, -4536, -4540, -4544)
'0070f546    8d45c0                  lea eax, dword ptr [ebp-40]
'0070f549    50                      push eax
'0070f54a    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0070f54d    51                      push ecx
'0070f54e    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070f550    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'0070f556    8d9530ffffff            lea edx, dword ptr [ebp+ffffff30]
'0070f55c    52                      push edx
'0070f55d    8d8540ffffff            lea eax, dword ptr [ebp+ffffff40]
'0070f563    50                      push eax
'0070f564    8d8d50ffffff            lea ecx, dword ptr [ebp+ffffff50]
'0070f56a    51                      push ecx
'0070f56b    8d9560ffffff            lea edx, dword ptr [ebp+ffffff60]
'0070f571    52                      push edx
'0070f572    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'0070f578    50                      push eax
'0070f579    8d4d80                  lea ecx, dword ptr [ebp-80]
'0070f57c    51                      push ecx
'0070f57d    8d5590                  lea edx, dword ptr [ebp-70]
'0070f580    52                      push edx
'0070f581    8d45a0                  lea eax, dword ptr [ebp-60]
'0070f584    50                      push eax
'0070f585    8d4db0                  lea ecx, dword ptr [ebp-50]
'0070f588    51                      push ecx
'0070f589    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0070f58b    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'0070f591    83c454                  add esp, 54
'0070f594    c3                      ret
'0070f595    c3                      ret
'0070f596    8b4508                  mov eax, dword ptr [ebp+08]
'0070f599    8b10                    mov edx, dword ptr [eax]
'0070f59b    50                      push eax
'0070f59c    ff5208                  call dword ptr [edx+08]
'0070f59f    8b45f4                  mov eax, dword ptr [ebp-0c]
'0070f5a2    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'0070f5a5    64890d00000000          mov dword ptr fs:[00000000], ecx
'0070f5ac    5f                      pop edi
'0070f5ad    5e                      pop esi
'0070f5ae    5b                      pop ebx
'0070f5af    8be5                    mov esp, ebp
'0070f5b1    5d                      pop ebp
'0070f5b2    c21000                  ret 0010


End Sub


Private Sub vspersonnage_Event39()
'0070f690    55                      push ebp
'0070f691    8bec                    mov ebp, esp
'0070f693    83ec14                  sub esp, 14

' *** Reference to "__vbaExceptHandler"
'0070f696    6866724000              push 00407266
'0070f69b    64a100000000            mov ax, word ptr fs:[00000000]
'0070f6a1    50                      push eax
    'Reference to '__except_list'
'0070f6a2    64892500000000          mov dword ptr fs:[00000000], esp
'0070f6a9    81ec2c010000            sub esp, 0000012c
'0070f6af    53                      push ebx
'0070f6b0    56                      push esi
'0070f6b1    57                      push edi
'0070f6b2    8965ec                  mov dword ptr [ebp-14], esp
'0070f6b5    c745f0f06c4000          mov dword ptr [ebp-10], 00406cf0
'0070f6bc    8b4508                  mov eax, dword ptr [ebp+08]
'0070f6bf    8bc8                    mov ecx, eax
'0070f6c1    83e101                  and ecx, 01
'0070f6c4    894df4                  mov dword ptr [ebp-0c], ecx
'0070f6c7    83e0fe                  and eax, fffffffe
'0070f6ca    894508                  mov dword ptr [ebp+08], eax
'0070f6cd    33f6                    xor esi, esi
'0070f6cf    8975f8                  mov dword ptr [ebp-08], esi
'0070f6d2    8b10                    mov edx, dword ptr [eax]
'0070f6d4    50                      push eax
'0070f6d5    ff5204                  call dword ptr [edx+04]
'0070f6d8    8975e0                  mov dword ptr [ebp-20], esi
'0070f6db    8975dc                  mov dword ptr [ebp-24], esi
'0070f6de    8975d8                  mov dword ptr [ebp-28], esi
'0070f6e1    8975d4                  mov dword ptr [ebp-2c], esi
'0070f6e4    8975d0                  mov dword ptr [ebp-30], esi
'0070f6e7    8975cc                  mov dword ptr [ebp-34], esi
'0070f6ea    8975c8                  mov dword ptr [ebp-38], esi
'0070f6ed    8975c4                  mov dword ptr [ebp-3c], esi
'0070f6f0    8975c0                  mov dword ptr [ebp-40], esi
'0070f6f3    8975b0                  mov dword ptr [ebp-50], esi
'0070f6f6    8975a0                  mov dword ptr [ebp-60], esi
'0070f6f9    897590                  mov dword ptr [ebp-70], esi
'0070f6fc    897580                  mov dword ptr [ebp-80], esi
'0070f6ff    89b570ffffff            mov dword ptr [ebp+ffffff70], esi
'0070f705    89b560ffffff            mov dword ptr [ebp+ffffff60], esi
'0070f70b    89b550ffffff            mov dword ptr [ebp+ffffff50], esi
'0070f711    89b540ffffff            mov dword ptr [ebp+ffffff40], esi
'0070f717    89b530ffffff            mov dword ptr [ebp+ffffff30], esi
'0070f71d    89b520ffffff            mov dword ptr [ebp+ffffff20], esi
'0070f723    89b510ffffff            mov dword ptr [ebp+ffffff10], esi
'0070f729    89b500ffffff            mov dword ptr [ebp+ffffff00], esi
'0070f72f    89b5f0feffff            mov dword ptr [ebp+fffffef0], esi
'0070f735    89b5e0feffff            mov dword ptr [ebp+fffffee0], esi
'0070f73b    89b5dcfeffff            mov dword ptr [ebp+fffffedc], esi
'0070f741    66393528b07200          cmp word ptr [0072b028], si
'0070f748    0f853e030000            jne 70fa8c
'0070f74e    6a01                    push 01

' *** Reference to "__vbaOnError"
'0070f750    ff15b0104000            call dword ptr [004010b0]
On Error Goto handler_0
'0070f756    e931030000              jmp 70fa8c

' *** Reference to "rtcErrObj"
'0070f75b    8b1d6c124000            mov ebx, dword ptr [0040126c]
'0070f761    ffd3                    call ebx
'0070f763    50                      push eax
'0070f764    8d45c4                  lea eax, dword ptr [ebp-3c]
'0070f767    50                      push eax

' *** Reference to "__vbaObjSet"
'0070f768    8b3db4104000            mov edi, dword ptr [004010b4]
'0070f76e    ffd7                    call edi
Set var_9 = Err
'0070f770    8bf0                    mov esi, eax
'0070f772    8b0e                    mov ecx, dword ptr [esi]
'0070f774    8d55e0                  lea edx, dword ptr [ebp-20]
'0070f777    52                      push edx
'0070f778    56                      push esi
'0070f779    ff512c                  call dword ptr [ecx+2c]
var_3 = var_9.Description
'0070f77c    dbe2                    fnclex
'0070f77e    85c0                    test ax, ax
'0070f780    7d0f                    jge 70f791
'0070f782    6a2c                    push 2c
'0070f784    685c084300              push 0043085c
'0070f789    56                      push esi
'0070f78a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070f78b    ff1580104000            call dword ptr [00401080]
'0070f791    ffd3                    call ebx
'0070f793    50                      push eax
'0070f794    8d45c0                  lea eax, dword ptr [ebp-40]
'0070f797    50                      push eax
'0070f798    ffd7                    call edi
Set var_5 = Err
'0070f79a    8bf0                    mov esi, eax
'0070f79c    8b0e                    mov ecx, dword ptr [esi]
'0070f79e    8d95dcfeffff            lea edx, dword ptr [ebp+fffffedc]
'0070f7a4    52                      push edx
'0070f7a5    56                      push esi
'0070f7a6    ff511c                  call dword ptr [ecx+1c]
var_10 = var_5.Number
'0070f7a9    dbe2                    fnclex
'0070f7ab    85c0                    test ax, ax
'0070f7ad    7d0f                    jge 70f7be
'0070f7af    6a1c                    push 1c
'0070f7b1    685c084300              push 0043085c
'0070f7b6    56                      push esi
'0070f7b7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070f7b8    ff1580104000            call dword ptr [00401080]
'0070f7be    b904000280              mov ecx, 80020004
'0070f7c3    894d98                  mov dword ptr [ebp-68], ecx
'0070f7c6    b80a000000              mov eax, 0000000a
'0070f7cb    894590                  mov dword ptr [ebp-70], eax
'0070f7ce    894da8                  mov dword ptr [ebp-58], ecx
'0070f7d1    8945a0                  mov dword ptr [ebp-60], eax
'0070f7d4    c78528ffffff24b07200    mov dword ptr [ebp+ffffff28], 0072b024
'0070f7de    c78520ffffff08400000    mov dword ptr [ebp+ffffff20], 00004008
'0070f7e8    6814084300              push 00430814
'0070f7ed    8b45e0                  mov eax, dword ptr [ebp-20]
'0070f7f0    50                      push eax

' *** Reference to "__vbaStrCat"
'0070f7f1    8b3570104000            mov esi, dword ptr [00401070]
'0070f7f7    ffd6                    call esi
var_63 = ("L'erreur suivante s'est produite : ") & (var_3)
'0070f7f9    8bd0                    mov edx, eax
'0070f7fb    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaStrMove"
'0070f7fe    8b3dd0124000            mov edi, dword ptr [004012d0]
'0070f804    ffd7                    call edi
'0070f806    50                      push eax
'0070f807    6870084300              push 00430870
'0070f80c    ffd6                    call esi
var_76 = (var_63) & (vbCrLf)
'0070f80e    8bd0                    mov edx, eax
'0070f810    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0070f813    ffd7                    call edi
'0070f815    50                      push eax
'0070f816    6870084300              push 00430870
'0070f81b    ffd6                    call esi
var_12 = (var_76) & (vbCrLf)
'0070f81d    8bd0                    mov edx, eax
'0070f81f    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0070f822    ffd7                    call edi
'0070f824    50                      push eax
'0070f825    6880084300              push 00430880
'0070f82a    ffd6                    call esi
var_13 = (var_12) & (" numéro d'erreur (")
'0070f82c    8bd0                    mov edx, eax
'0070f82e    8d4dd0                  lea ecx, dword ptr [ebp-30]
'0070f831    ffd7                    call edi
'0070f833    50                      push eax
'0070f834    8b8ddcfeffff            mov ecx, dword ptr [ebp+fffffedc]
'0070f83a    51                      push ecx

' *** Reference to "__vbaStrI4"
'0070f83b    ff1520104000            call dword ptr [00401020]
'0070f841    8bd0                    mov edx, eax
'0070f843    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070f846    ffd7                    call edi
'0070f848    50                      push eax
'0070f849    ffd6                    call esi
var_127 = (var_13) & (CStr(var_10))
'0070f84b    8bd0                    mov edx, eax
'0070f84d    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0070f850    ffd7                    call edi
'0070f852    50                      push eax
'0070f853    68ac084300              push 004308ac
'0070f858    ffd6                    call esi
var_15 = (var_127) & (" )")
'0070f85a    8945b8                  mov dword ptr [ebp-48], eax
'0070f85d    bf08000000              mov edi, 00000008
'0070f862    897db0                  mov dword ptr [ebp-50], edi
'0070f865    8d5590                  lea edx, dword ptr [ebp-70]
'0070f868    52                      push edx
'0070f869    8d45a0                  lea eax, dword ptr [ebp-60]
'0070f86c    50                      push eax
'0070f86d    8d8d20ffffff            lea ecx, dword ptr [ebp+ffffff20]
'0070f873    51                      push ecx
'0070f874    6a10                    push 10
'0070f876    8d55b0                  lea edx, dword ptr [ebp-50]
'0070f879    52                      push edx

' *** Reference to "rtcMsgBox"
'0070f87a    ff15b8104000            call dword ptr [004010b8]
var_128 = MsgBox(var_15, 16, vbNullString)
'0070f880    8d45c8                  lea eax, dword ptr [ebp-38]
'0070f883    50                      push eax
'0070f884    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070f887    51                      push ecx
'0070f888    8d55d0                  lea edx, dword ptr [ebp-30]
'0070f88b    52                      push edx
'0070f88c    8d45d4                  lea eax, dword ptr [ebp-2c]
'0070f88f    50                      push eax
'0070f890    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0070f893    51                      push ecx
'0070f894    8d55dc                  lea edx, dword ptr [ebp-24]
'0070f897    52                      push edx
'0070f898    8d45e0                  lea eax, dword ptr [ebp-20]
'0070f89b    50                      push eax
'0070f89c    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0070f89e    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4508, -4512, -4516, -4520, -4524, -4528)
'0070f8a4    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0070f8a7    51                      push ecx
'0070f8a8    8d55c4                  lea edx, dword ptr [ebp-3c]
'0070f8ab    52                      push edx
'0070f8ac    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070f8ae    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'0070f8b4    8d4590                  lea eax, dword ptr [ebp-70]
'0070f8b7    50                      push eax
'0070f8b8    8d4da0                  lea ecx, dword ptr [ebp-60]
'0070f8bb    51                      push ecx
'0070f8bc    8d55b0                  lea edx, dword ptr [ebp-50]
'0070f8bf    52                      push edx
'0070f8c0    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'0070f8c2    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8)
'0070f8c8    83c43c                  add esp, 3c
'0070f8cb    8d45b0                  lea eax, dword ptr [ebp-50]
'0070f8ce    50                      push eax

' *** Reference to "rtcGetPresentDate"
'0070f8cf    ff15f4124000            call dword ptr [004012f4]
var_15 = Now()
'0070f8d5    c78528ffffffb8084300    mov dword ptr [ebp+ffffff28], 004308b8
'0070f8df    89bd20ffffff            mov dword ptr [ebp+ffffff20], edi
'0070f8e5    8d9520ffffff            lea edx, dword ptr [ebp+ffffff20]
'0070f8eb    8d4da0                  lea ecx, dword ptr [ebp-60]

' *** Reference to "__vbaVarDup"
'0070f8ee    ff1598124000            call dword ptr [00401298]
var_7 = ("dd/MM/yyyy hh:mm:ss")
'0070f8f4    6a01                    push 01
'0070f8f6    6a01                    push 01
'0070f8f8    8d4da0                  lea ecx, dword ptr [ebp-60]
'0070f8fb    51                      push ecx
'0070f8fc    8d55b0                  lea edx, dword ptr [ebp-50]
'0070f8ff    52                      push edx
'0070f900    8d4590                  lea eax, dword ptr [ebp-70]
'0070f903    50                      push eax

' *** Reference to "rtcVarFromFormatVar"
'0070f904    ff1574104000            call dword ptr [00401074]
'0070f90a    ffd3                    call ebx
'0070f90c    50                      push eax
'0070f90d    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0070f910    51                      push ecx

' *** Reference to "__vbaObjSet"
'0070f911    ff15b4104000            call dword ptr [004010b4]
Set var_9 = Err
'0070f917    8bf0                    mov esi, eax
'0070f919    8b16                    mov edx, dword ptr [esi]
'0070f91b    8d85dcfeffff            lea eax, dword ptr [ebp+fffffedc]
'0070f921    50                      push eax
'0070f922    56                      push esi
'0070f923    ff521c                  call dword ptr [edx+1c]
var_10 = var_9.Number
'0070f926    dbe2                    fnclex
'0070f928    85c0                    test ax, ax
'0070f92a    7d0f                    jge 70f93b
'0070f92c    6a1c                    push 1c
'0070f92e    685c084300              push 0043085c
'0070f933    56                      push esi
'0070f934    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070f935    ff1580104000            call dword ptr [00401080]
'0070f93b    ffd3                    call ebx
'0070f93d    50                      push eax
'0070f93e    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0070f941    51                      push ecx

' *** Reference to "__vbaObjSet"
'0070f942    ff15b4104000            call dword ptr [004010b4]
Set var_5 = Err
'0070f948    8bf0                    mov esi, eax
'0070f94a    8b16                    mov edx, dword ptr [esi]
'0070f94c    8d45e0                  lea eax, dword ptr [ebp-20]
'0070f94f    50                      push eax
'0070f950    56                      push esi
'0070f951    ff522c                  call dword ptr [edx+2c]
var_3 = var_5.Description
'0070f954    dbe2                    fnclex
'0070f956    85c0                    test ax, ax
'0070f958    7d0f                    jge 70f969
'0070f95a    6a2c                    push 2c
'0070f95c    685c084300              push 0043085c
'0070f961    56                      push esi
'0070f962    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070f963    ff1580104000            call dword ptr [00401080]
'0070f969    c78518ffffffe4084300    mov dword ptr [ebp+ffffff18], 004308e4
'0070f973    89bd10ffffff            mov dword ptr [ebp+ffffff10], edi
'0070f979    8b8ddcfeffff            mov ecx, dword ptr [ebp+fffffedc]
'0070f97f    898d08ffffff            mov dword ptr [ebp+ffffff08], ecx
'0070f985    c78500ffffff03000000    mov dword ptr [ebp+ffffff00], 00000003
'0070f98f    c785f8feffff08094300    mov dword ptr [ebp+fffffef8], 00430908
'0070f999    89bdf0feffff            mov dword ptr [ebp+fffffef0], edi
'0070f99f    8b45e0                  mov eax, dword ptr [ebp-20]
'0070f9a2    c745e000000000          mov dword ptr [ebp-20], 00000000
'0070f9a9    898558ffffff            mov dword ptr [ebp+ffffff58], eax
'0070f9af    89bd50ffffff            mov dword ptr [ebp+ffffff50], edi
'0070f9b5    c785e8feffff38154500    mov dword ptr [ebp+fffffee8], 00451538
'0070f9bf    89bde0feffff            mov dword ptr [ebp+fffffee0], edi
'0070f9c5    8d5590                  lea edx, dword ptr [ebp-70]
'0070f9c8    52                      push edx
'0070f9c9    8d8510ffffff            lea eax, dword ptr [ebp+ffffff10]
'0070f9cf    50                      push eax
'0070f9d0    8d4d80                  lea ecx, dword ptr [ebp-80]
'0070f9d3    51                      push ecx

' *** Reference to "__vbaVarCat"
'0070f9d4    8b3508124000            mov esi, dword ptr [00401208]
'0070f9da    ffd6                    call esi
'0070f9dc    50                      push eax
'0070f9dd    8d9500ffffff            lea edx, dword ptr [ebp+ffffff00]
'0070f9e3    52                      push edx
'0070f9e4    8d8570ffffff            lea eax, dword ptr [ebp+ffffff70]
'0070f9ea    50                      push eax
'0070f9eb    ffd6                    call esi
'0070f9ed    50                      push eax
'0070f9ee    8d8df0feffff            lea ecx, dword ptr [ebp+fffffef0]
'0070f9f4    51                      push ecx
'0070f9f5    8d9560ffffff            lea edx, dword ptr [ebp+ffffff60]
'0070f9fb    52                      push edx
'0070f9fc    ffd6                    call esi
'0070f9fe    50                      push eax
'0070f9ff    8d8550ffffff            lea eax, dword ptr [ebp+ffffff50]
'0070fa05    50                      push eax
'0070fa06    8d8d40ffffff            lea ecx, dword ptr [ebp+ffffff40]
'0070fa0c    51                      push ecx
'0070fa0d    ffd6                    call esi
'0070fa0f    50                      push eax
'0070fa10    8d95e0feffff            lea edx, dword ptr [ebp+fffffee0]
'0070fa16    52                      push edx
'0070fa17    8d8530ffffff            lea eax, dword ptr [ebp+ffffff30]
'0070fa1d    50                      push eax
'0070fa1e    ffd6                    call esi
'0070fa20    50                      push eax
'0070fa21    33c9                    xor ecx, ecx
'0070fa23    668b0d2ab07200          mov cx, word ptr [0072b02a]
'0070fa2a    51                      push ecx
'0070fa2b    6884094300              push 00430984

' *** Reference to "__vbaPrintFile"
'0070fa30    ff15b8114000            call dword ptr [004011b8]
Print #0, 
'0070fa36    8d55c0                  lea edx, dword ptr [ebp-40]
'0070fa39    52                      push edx
'0070fa3a    8d45c4                  lea eax, dword ptr [ebp-3c]
'0070fa3d    50                      push eax
'0070fa3e    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070fa40    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'0070fa46    8d8d30ffffff            lea ecx, dword ptr [ebp+ffffff30]
'0070fa4c    51                      push ecx
'0070fa4d    8d9540ffffff            lea edx, dword ptr [ebp+ffffff40]
'0070fa53    52                      push edx
'0070fa54    8d8550ffffff            lea eax, dword ptr [ebp+ffffff50]
'0070fa5a    50                      push eax
'0070fa5b    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'0070fa61    51                      push ecx
'0070fa62    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0070fa68    52                      push edx
'0070fa69    8d4580                  lea eax, dword ptr [ebp-80]
'0070fa6c    50                      push eax
'0070fa6d    8d4d90                  lea ecx, dword ptr [ebp-70]
'0070fa70    51                      push ecx
'0070fa71    8d55a0                  lea edx, dword ptr [ebp-60]
'0070fa74    52                      push edx
'0070fa75    8d45b0                  lea eax, dword ptr [ebp-50]
'0070fa78    50                      push eax
'0070fa79    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0070fa7b    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'0070fa81    83c440                  add esp, 40
'0070fa84    6a00                    push 00

' *** Reference to "__vbaResume"
'0070fa86    ff1568104000            call dword ptr [00401068]

' *** Reference to "__vbaExitProc"
'0070fa8c    ff15a0104000            call dword ptr [004010a0]
'0070fa92    680dfb7000              push 0070fb0d
'0070fa97    eb73                    jmp 70fb0c
Resume handler_70FB0C
'0070fa99    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0070fa9c    51                      push ecx
'0070fa9d    8d55cc                  lea edx, dword ptr [ebp-34]
'0070faa0    52                      push edx
'0070faa1    8d45d0                  lea eax, dword ptr [ebp-30]
'0070faa4    50                      push eax
'0070faa5    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0070faa8    51                      push ecx
'0070faa9    8d55d8                  lea edx, dword ptr [ebp-28]
'0070faac    52                      push edx
'0070faad    8d45dc                  lea eax, dword ptr [ebp-24]
'0070fab0    50                      push eax
'0070fab1    8d4de0                  lea ecx, dword ptr [ebp-20]
'0070fab4    51                      push ecx
'0070fab5    6a07                    push 07

' *** Reference to "__vbaFreeStrList"
'0070fab7    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_3, -4508, -4512, -4516, -4520, -4524, -4528)
'0070fabd    8d55c0                  lea edx, dword ptr [ebp-40]
'0070fac0    52                      push edx
'0070fac1    8d45c4                  lea eax, dword ptr [ebp-3c]
'0070fac4    50                      push eax
'0070fac5    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'0070fac7    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_9, var_5)
'0070facd    8d8d30ffffff            lea ecx, dword ptr [ebp+ffffff30]
'0070fad3    51                      push ecx
'0070fad4    8d9540ffffff            lea edx, dword ptr [ebp+ffffff40]
'0070fada    52                      push edx
'0070fadb    8d8550ffffff            lea eax, dword ptr [ebp+ffffff50]
'0070fae1    50                      push eax
'0070fae2    8d8d60ffffff            lea ecx, dword ptr [ebp+ffffff60]
'0070fae8    51                      push ecx
'0070fae9    8d9570ffffff            lea edx, dword ptr [ebp+ffffff70]
'0070faef    52                      push edx
'0070faf0    8d4580                  lea eax, dword ptr [ebp-80]
'0070faf3    50                      push eax
'0070faf4    8d4d90                  lea ecx, dword ptr [ebp-70]
'0070faf7    51                      push ecx
'0070faf8    8d55a0                  lea edx, dword ptr [ebp-60]
'0070fafb    52                      push edx
'0070fafc    8d45b0                  lea eax, dword ptr [ebp-50]
'0070faff    50                      push eax
'0070fb00    6a09                    push 09

' *** Reference to "__vbaFreeVarList"
'0070fb02    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_6, var_7, var_8, var_18, var_19, var_20, var_21, var_22, var_23)
'0070fb08    83c454                  add esp, 54
'0070fb0b    c3                      ret
'0070fb0c    c3                      ret
'0070fb0d    8b4508                  mov eax, dword ptr [ebp+08]
'0070fb10    8b08                    mov ecx, dword ptr [eax]
'0070fb12    50                      push eax
'0070fb13    ff5108                  call dword ptr [ecx+08]
'0070fb16    8b45f4                  mov eax, dword ptr [ebp-0c]
'0070fb19    8b4de4                  mov ecx, dword ptr [ebp-1c]
    'Reference to '__except_list'
'0070fb1c    64890d00000000          mov dword ptr fs:[00000000], ecx
'0070fb23    5f                      pop edi
'0070fb24    5e                      pop esi
'0070fb25    5b                      pop ebx
'0070fb26    8be5                    mov esp, ebp
'0070fb28    5d                      pop ebp
'0070fb29    c21000                  ret 0010


End Sub


'Event for btnEnregistrer
Private Sub btnEnregistrer_Click()
'0070cca0    55                      push ebp
'0070cca1    8bec                    mov ebp, esp
'0070cca3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'0070cca6    6866724000              push 00407266
'0070ccab    64a100000000            mov ax, word ptr fs:[00000000]
'0070ccb1    50                      push eax
    'Reference to '__except_list'
'0070ccb2    64892500000000          mov dword ptr fs:[00000000], esp
'0070ccb9    81ec54020000            sub esp, 00000254
'0070ccbf    53                      push ebx
'0070ccc0    56                      push esi
'0070ccc1    57                      push edi
'0070ccc2    8965f4                  mov dword ptr [ebp-0c], esp
'0070ccc5    c745f8406c4000          mov dword ptr [ebp-08], 00406c40
'0070cccc    8b7508                  mov esi, dword ptr [ebp+08]
'0070cccf    8bc6                    mov eax, esi
'0070ccd1    83e001                  and eax, 01
'0070ccd4    8945fc                  mov dword ptr [ebp-04], eax
'0070ccd7    83e6fe                  and esi, fffffffe
'0070ccda    8b0e                    mov ecx, dword ptr [esi]
'0070ccdc    56                      push esi
'0070ccdd    897508                  mov dword ptr [ebp+08], esi
'0070cce0    ff5104                  call dword ptr [ecx+04]
'0070cce3    33db                    xor ebx, ebx
'0070cce5    8d7dcc                  lea edi, dword ptr [ebp-34]
'0070cce8    57                      push edi
'0070cce9    8b1548b07200            mov edx, dword ptr [0072b048]
'0070ccef    83ec10                  sub esp, 10
'0070ccf2    b90a000000              mov ecx, 0000000a
'0070ccf7    898d1cffffff            mov dword ptr [ebp+ffffff1c], ecx
'0070ccfd    8bfc                    mov edi, esp
'0070ccff    890f                    mov dword ptr [edi], ecx
'0070cd01    8b8d10ffffff            mov ecx, dword ptr [ebp+ffffff10]
'0070cd07    894f04                  mov dword ptr [edi+04], ecx
'0070cd0a    b804000280              mov eax, 80020004
'0070cd0f    894708                  mov dword ptr [edi+08], eax
'0070cd12    898514ffffff            mov dword ptr [ebp+ffffff14], eax
'0070cd18    898524ffffff            mov dword ptr [ebp+ffffff24], eax
'0070cd1e    8b8518ffffff            mov eax, dword ptr [ebp+ffffff18]
'0070cd24    89470c                  mov dword ptr [edi+0c], eax
'0070cd27    8b851cffffff            mov eax, dword ptr [ebp+ffffff1c]
'0070cd2d    83ec10                  sub esp, 10
'0070cd30    8bcc                    mov ecx, esp
'0070cd32    8901                    mov dword ptr [ecx], eax
'0070cd34    8b8520ffffff            mov eax, dword ptr [ebp+ffffff20]
'0070cd3a    894104                  mov dword ptr [ecx+04], eax
'0070cd3d    8b8524ffffff            mov eax, dword ptr [ebp+ffffff24]
'0070cd43    894108                  mov dword ptr [ecx+08], eax
'0070cd46    8b8528ffffff            mov eax, dword ptr [ebp+ffffff28]
'0070cd4c    89410c                  mov dword ptr [ecx+0c], eax
'0070cd4f    83ec10                  sub esp, 10
'0070cd52    b803000000              mov eax, 00000003
'0070cd57    8bcc                    mov ecx, esp
'0070cd59    8901                    mov dword ptr [ecx], eax
'0070cd5b    8b8530ffffff            mov eax, dword ptr [ebp+ffffff30]
'0070cd61    895de4                  mov dword ptr [ebp-1c], ebx
'0070cd64    895de0                  mov dword ptr [ebp-20], ebx
'0070cd67    895ddc                  mov dword ptr [ebp-24], ebx
'0070cd6a    895dd8                  mov dword ptr [ebp-28], ebx
'0070cd6d    895dd4                  mov dword ptr [ebp-2c], ebx
'0070cd70    895dd0                  mov dword ptr [ebp-30], ebx
'0070cd73    895dcc                  mov dword ptr [ebp-34], ebx
'0070cd76    895dc8                  mov dword ptr [ebp-38], ebx
'0070cd79    895dc4                  mov dword ptr [ebp-3c], ebx
'0070cd7c    895dc0                  mov dword ptr [ebp-40], ebx
'0070cd7f    895dbc                  mov dword ptr [ebp-44], ebx
'0070cd82    895dac                  mov dword ptr [ebp-54], ebx
'0070cd85    895d9c                  mov dword ptr [ebp-64], ebx
'0070cd88    895d8c                  mov dword ptr [ebp-74], ebx
'0070cd8b    899d7cffffff            mov dword ptr [ebp+ffffff7c], ebx
'0070cd91    899d6cffffff            mov dword ptr [ebp+ffffff6c], ebx
'0070cd97    899d5cffffff            mov dword ptr [ebp+ffffff5c], ebx
'0070cd9d    899d4cffffff            mov dword ptr [ebp+ffffff4c], ebx
'0070cda3    899d3cffffff            mov dword ptr [ebp+ffffff3c], ebx
'0070cda9    899dccfeffff            mov dword ptr [ebp+fffffecc], ebx
'0070cdaf    899d6cfeffff            mov dword ptr [ebp+fffffe6c], ebx
'0070cdb5    899d5cfeffff            mov dword ptr [ebp+fffffe5c], ebx
'0070cdbb    899d4cfeffff            mov dword ptr [ebp+fffffe4c], ebx
'0070cdc1    899d3cfeffff            mov dword ptr [ebp+fffffe3c], ebx
'0070cdc7    8b12                    mov edx, dword ptr [edx]
'0070cdc9    899decfeffff            mov dword ptr [ebp+fffffeec], ebx
'0070cdcf    899dbcfeffff            mov dword ptr [ebp+fffffebc], ebx
'0070cdd5    899dacfeffff            mov dword ptr [ebp+fffffeac], ebx
'0070cddb    899d9cfeffff            mov dword ptr [ebp+fffffe9c], ebx
'0070cde1    899d8cfeffff            mov dword ptr [ebp+fffffe8c], ebx
'0070cde7    899d7cfeffff            mov dword ptr [ebp+fffffe7c], ebx
'0070cded    899d2cfeffff            mov dword ptr [ebp+fffffe2c], ebx
'0070cdf3    899d1cfeffff            mov dword ptr [ebp+fffffe1c], ebx
'0070cdf9    c78534ffffff01000000    mov dword ptr [ebp+ffffff34], 00000001
'0070ce03    894104                  mov dword ptr [ecx+04], eax
'0070ce06    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'0070ce0c    894108                  mov dword ptr [ecx+08], eax
'0070ce0f    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'0070ce15    89410c                  mov dword ptr [ecx+0c], eax
'0070ce18    8b0d48b07200            mov ecx, dword ptr [0072b048]
'0070ce1e    68c8b14300              push 0043b1c8
'0070ce23    51                      push ecx
'0070ce24    ff92bc000000            call dword ptr [edx+000000bc]
'0070ce2a    dbe2                    fnclex
'0070ce2c    3bc3                    cmp eax, ebx
'0070ce2e    7d18                    jge 70ce48

If (0 < 0) Then
'0070ce30    8b1548b07200            mov edx, dword ptr [0072b048]
'0070ce36    68bc000000              push 000000bc
'0070ce3b    68301f4300              push 00431f30
'0070ce40    52                      push edx
'0070ce41    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070ce42    ff1580104000            call dword ptr [00401080]
    
End If
'0070ce48    8b45cc                  mov eax, dword ptr [ebp-34]

' *** Reference to "__vbaObjSet"
'0070ce4b    8b3db4104000            mov edi, dword ptr [004010b4]
'0070ce51    50                      push eax
'0070ce52    8d45e4                  lea eax, dword ptr [ebp-1c]
'0070ce55    50                      push eax
'0070ce56    895dcc                  mov dword ptr [ebp-34], ebx
'0070ce59    ffd7                    call edi
Set var_40 = Nothing
'0070ce5b    8b45e4                  mov eax, dword ptr [ebp-1c]
'0070ce5e    8b08                    mov ecx, dword ptr [eax]
'0070ce60    68b4d44400              push 0044d4b4
'0070ce65    50                      push eax
'0070ce66    ff5144                  call dword ptr [ecx+44]
'0070ce69    dbe2                    fnclex
'0070ce6b    3bc3                    cmp eax, ebx
'0070ce6d    7d12                    jge 70ce81

If (var_40 < 0) Then
'0070ce6f    8b55e4                  mov edx, dword ptr [ebp-1c]
'0070ce72    6a44                    push 44
'0070ce74    6830314300              push 00433130
'0070ce79    52                      push edx
'0070ce7a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070ce7b    ff1580104000            call dword ptr [00401080]
    
End If
'0070ce81    8b06                    mov eax, dword ptr [esi]
'0070ce83    53                      push ebx
'0070ce84    6a07                    push 07
'0070ce86    56                      push esi
'0070ce87    ff9020030000            call dword ptr [eax+00000320]
'0070ce8d    50                      push eax
'0070ce8e    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070ce91    51                      push ecx
'0070ce92    ffd7                    call edi
Set var_43 = Nothing
'0070ce94    50                      push eax
'0070ce95    8d55ac                  lea edx, dword ptr [ebp-54]
'0070ce98    52                      push edx

' *** Reference to "__vbaLateIdCallLd"
'0070ce99    ff157c114000            call dword ptr [0040117c]
var_50 = var_43.UNK_-256 - 12_7
'0070ce9f    83c410                  add esp, 10
'0070cea2    50                      push eax

' *** Reference to "__vbaI4Var"
'0070cea3    ff157c124000            call dword ptr [0040127c]
'0070cea9    8bc8                    mov ecx, eax
'0070ceab    83e901                  sub ecx, 01
var_num3 = CLng(var_50) - 1
'0070ceae    0f8050140000            jo 70e304

' *** Reference to "__vbaI2I4"
'0070ceb4    ff1550114000            call dword ptr [00401150]
'0070ceba    bb01000000              mov ebx, 00000001
'0070cebf    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070cec2    895de8                  mov dword ptr [ebp-18], ebx
'0070cec5    8985bcfdffff            mov dword ptr [ebp+fffffdbc], eax

' *** Reference to "__vbaFreeObj"
'0070cecb    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'0070ced1    8d4dac                  lea ecx, dword ptr [ebp-54]

' *** Reference to "__vbaFreeVar"
'0070ced4    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_50)
'0070ceda    663b9dbcfdffff          cmp bx, word ptr [ebp+fffffdbc]
'0070cee1    0f8f46130000            jg 70e22d

Do While (1 <= WORD PTR [EBP+FFFFFDBC])
'0070cee7    83ec10                  sub esp, 10
'0070ceea    8bd4                    mov edx, esp
'0070ceec    33c0                    xor eax, eax
    var_num1 = Empty
'0070ceee    8985f4feffff            mov dword ptr [ebp+fffffef4], eax
'0070cef4    b903000000              mov ecx, 00000003
'0070cef9    890a                    mov dword ptr [edx], ecx
'0070cefb    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'0070cf01    894a04                  mov dword ptr [edx+04], ecx
'0070cf04    894208                  mov dword ptr [edx+08], eax
'0070cf07    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'0070cf0d    89420c                  mov dword ptr [edx+0c], eax
'0070cf10    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'0070cf16    83ec10                  sub esp, 10
'0070cf19    8bcc                    mov ecx, esp
'0070cf1b    b802000000              mov eax, 00000002
'0070cf20    8901                    mov dword ptr [ecx], eax
'0070cf22    895104                  mov dword ptr [ecx+04], edx
'0070cf25    8b9518ffffff            mov edx, dword ptr [ebp+ffffff18]
'0070cf2b    66899d14ffffff          mov word ptr [ebp+ffffff14], bx
'0070cf32    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'0070cf38    894108                  mov dword ptr [ecx+08], eax
'0070cf3b    89510c                  mov dword ptr [ecx+0c], edx
'0070cf3e    8b95f0feffff            mov edx, dword ptr [ebp+fffffef0]
'0070cf44    83ec10                  sub esp, 10
'0070cf47    8bc4                    mov eax, esp
'0070cf49    c785ecfeffff02000000    mov dword ptr [ebp+fffffeec], 00000002
'0070cf53    8b8decfeffff            mov ecx, dword ptr [ebp+fffffeec]
'0070cf59    8908                    mov dword ptr [eax], ecx
'0070cf5b    8b8df4feffff            mov ecx, dword ptr [ebp+fffffef4]
'0070cf61    895004                  mov dword ptr [eax+04], edx
'0070cf64    8b95f8feffff            mov edx, dword ptr [ebp+fffffef8]
'0070cf6a    6a03                    push 03
'0070cf6c    894808                  mov dword ptr [eax+08], ecx
'0070cf6f    689e000000              push 0000009e
'0070cf74    89500c                  mov dword ptr [eax+0c], edx
'0070cf77    8b06                    mov eax, dword ptr [esi]
'0070cf79    56                      push esi
'0070cf7a    ff9020030000            call dword ptr [eax+00000320]
'0070cf80    50                      push eax
'0070cf81    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070cf84    51                      push ecx
'0070cf85    ffd7                    call edi
    Set var_43 = Nothing
'0070cf87    50                      push eax
'0070cf88    8d55ac                  lea edx, dword ptr [ebp-54]
'0070cf8b    52                      push edx

' *** Reference to "__vbaLateIdCallLd"
'0070cf8c    ff157c114000            call dword ptr [0040117c]
    var_50 = var_43.UNK_-256 - 12_158
'0070cf92    b904000280              mov ecx, 80020004
'0070cf97    898d44feffff            mov dword ptr [ebp+fffffe44], ecx
'0070cf9d    898d54feffff            mov dword ptr [ebp+fffffe54], ecx
'0070cfa3    898d64feffff            mov dword ptr [ebp+fffffe64], ecx
'0070cfa9    898d74feffff            mov dword ptr [ebp+fffffe74], ecx
'0070cfaf    898dd4feffff            mov dword ptr [ebp+fffffed4], ecx
'0070cfb5    898da4feffff            mov dword ptr [ebp+fffffea4], ecx
'0070cfbb    898db4feffff            mov dword ptr [ebp+fffffeb4], ecx
'0070cfc1    8b4de4                  mov ecx, dword ptr [ebp-1c]
'0070cfc4    b80a000000              mov eax, 0000000a
'0070cfc9    89853cfeffff            mov dword ptr [ebp+fffffe3c], eax
'0070cfcf    89854cfeffff            mov dword ptr [ebp+fffffe4c], eax
'0070cfd5    89855cfeffff            mov dword ptr [ebp+fffffe5c], eax
'0070cfdb    89856cfeffff            mov dword ptr [ebp+fffffe6c], eax
'0070cfe1    8985ccfeffff            mov dword ptr [ebp+fffffecc], eax
'0070cfe7    8b11                    mov edx, dword ptr [ecx]
'0070cfe9    83c430                  add esp, 30
'0070cfec    8bcc                    mov ecx, esp
'0070cfee    8901                    mov dword ptr [ecx], eax
'0070cff0    8b8520feffff            mov eax, dword ptr [ebp+fffffe20]
'0070cff6    894104                  mov dword ptr [ecx+04], eax
'0070cff9    b804000280              mov eax, 80020004
'0070cffe    894108                  mov dword ptr [ecx+08], eax
'0070d001    8b8528feffff            mov eax, dword ptr [ebp+fffffe28]
'0070d007    89410c                  mov dword ptr [ecx+0c], eax
'0070d00a    83ec10                  sub esp, 10
'0070d00d    b80a000000              mov eax, 0000000a
'0070d012    8bcc                    mov ecx, esp
'0070d014    8901                    mov dword ptr [ecx], eax
'0070d016    8b8530feffff            mov eax, dword ptr [ebp+fffffe30]
'0070d01c    894104                  mov dword ptr [ecx+04], eax
'0070d01f    b804000280              mov eax, 80020004
'0070d024    894108                  mov dword ptr [ecx+08], eax
'0070d027    8b8538feffff            mov eax, dword ptr [ebp+fffffe38]
'0070d02d    89410c                  mov dword ptr [ecx+0c], eax
'0070d030    8b853cfeffff            mov eax, dword ptr [ebp+fffffe3c]
'0070d036    83ec10                  sub esp, 10
'0070d039    8bcc                    mov ecx, esp
'0070d03b    8901                    mov dword ptr [ecx], eax
'0070d03d    8b8540feffff            mov eax, dword ptr [ebp+fffffe40]
'0070d043    894104                  mov dword ptr [ecx+04], eax
'0070d046    8b8544feffff            mov eax, dword ptr [ebp+fffffe44]
'0070d04c    894108                  mov dword ptr [ecx+08], eax
'0070d04f    8b8548feffff            mov eax, dword ptr [ebp+fffffe48]
'0070d055    89410c                  mov dword ptr [ecx+0c], eax
'0070d058    8b854cfeffff            mov eax, dword ptr [ebp+fffffe4c]
'0070d05e    83ec10                  sub esp, 10
'0070d061    8bcc                    mov ecx, esp
'0070d063    8901                    mov dword ptr [ecx], eax
'0070d065    8b8550feffff            mov eax, dword ptr [ebp+fffffe50]
'0070d06b    894104                  mov dword ptr [ecx+04], eax
'0070d06e    8b8554feffff            mov eax, dword ptr [ebp+fffffe54]
'0070d074    894108                  mov dword ptr [ecx+08], eax
'0070d077    8b8558feffff            mov eax, dword ptr [ebp+fffffe58]
'0070d07d    89410c                  mov dword ptr [ecx+0c], eax
'0070d080    8b855cfeffff            mov eax, dword ptr [ebp+fffffe5c]
'0070d086    83ec10                  sub esp, 10
'0070d089    8bcc                    mov ecx, esp
'0070d08b    8901                    mov dword ptr [ecx], eax
'0070d08d    8b8560feffff            mov eax, dword ptr [ebp+fffffe60]
'0070d093    894104                  mov dword ptr [ecx+04], eax
'0070d096    8b8564feffff            mov eax, dword ptr [ebp+fffffe64]
'0070d09c    894108                  mov dword ptr [ecx+08], eax
'0070d09f    8b8568feffff            mov eax, dword ptr [ebp+fffffe68]
'0070d0a5    89410c                  mov dword ptr [ecx+0c], eax
'0070d0a8    8b856cfeffff            mov eax, dword ptr [ebp+fffffe6c]
'0070d0ae    83ec10                  sub esp, 10
'0070d0b1    8bcc                    mov ecx, esp
'0070d0b3    8901                    mov dword ptr [ecx], eax
'0070d0b5    8b8570feffff            mov eax, dword ptr [ebp+fffffe70]
'0070d0bb    894104                  mov dword ptr [ecx+04], eax
'0070d0be    8b8574feffff            mov eax, dword ptr [ebp+fffffe74]
'0070d0c4    894108                  mov dword ptr [ecx+08], eax
'0070d0c7    8b8578feffff            mov eax, dword ptr [ebp+fffffe78]
'0070d0cd    89410c                  mov dword ptr [ecx+0c], eax
'0070d0d0    83ec10                  sub esp, 10
'0070d0d3    8bcc                    mov ecx, esp
'0070d0d5    b80a000000              mov eax, 0000000a
'0070d0da    8901                    mov dword ptr [ecx], eax
'0070d0dc    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'0070d0e2    894104                  mov dword ptr [ecx+04], eax
'0070d0e5    b804000280              mov eax, 80020004
'0070d0ea    894108                  mov dword ptr [ecx+08], eax
'0070d0ed    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'0070d0f3    89410c                  mov dword ptr [ecx+0c], eax
'0070d0f6    83ec10                  sub esp, 10
'0070d0f9    8bcc                    mov ecx, esp
'0070d0fb    b80a000000              mov eax, 0000000a
'0070d100    8901                    mov dword ptr [ecx], eax
'0070d102    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'0070d108    894104                  mov dword ptr [ecx+04], eax
'0070d10b    b804000280              mov eax, 80020004
'0070d110    894108                  mov dword ptr [ecx+08], eax
'0070d113    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0070d119    89410c                  mov dword ptr [ecx+0c], eax
'0070d11c    83ec10                  sub esp, 10
'0070d11f    8bcc                    mov ecx, esp
'0070d121    b80a000000              mov eax, 0000000a
'0070d126    8901                    mov dword ptr [ecx], eax
'0070d128    8b85a0feffff            mov eax, dword ptr [ebp+fffffea0]
'0070d12e    894104                  mov dword ptr [ecx+04], eax
'0070d131    8b85a4feffff            mov eax, dword ptr [ebp+fffffea4]
'0070d137    894108                  mov dword ptr [ecx+08], eax
'0070d13a    8b85a8feffff            mov eax, dword ptr [ebp+fffffea8]
'0070d140    89410c                  mov dword ptr [ecx+0c], eax
'0070d143    83ec10                  sub esp, 10
'0070d146    8bcc                    mov ecx, esp
'0070d148    b80a000000              mov eax, 0000000a
'0070d14d    8901                    mov dword ptr [ecx], eax
'0070d14f    8b85b0feffff            mov eax, dword ptr [ebp+fffffeb0]
'0070d155    894104                  mov dword ptr [ecx+04], eax
'0070d158    8b85b4feffff            mov eax, dword ptr [ebp+fffffeb4]
'0070d15e    894108                  mov dword ptr [ecx+08], eax
'0070d161    8b85b8feffff            mov eax, dword ptr [ebp+fffffeb8]
'0070d167    89410c                  mov dword ptr [ecx+0c], eax
'0070d16a    83ec10                  sub esp, 10
'0070d16d    8bcc                    mov ecx, esp
'0070d16f    b80a000000              mov eax, 0000000a
'0070d174    8901                    mov dword ptr [ecx], eax
'0070d176    8b85c0feffff            mov eax, dword ptr [ebp+fffffec0]
'0070d17c    894104                  mov dword ptr [ecx+04], eax
'0070d17f    b804000280              mov eax, 80020004
'0070d184    894108                  mov dword ptr [ecx+08], eax
'0070d187    8b85c8feffff            mov eax, dword ptr [ebp+fffffec8]
'0070d18d    89410c                  mov dword ptr [ecx+0c], eax
'0070d190    8b85ccfeffff            mov eax, dword ptr [ebp+fffffecc]
'0070d196    83ec10                  sub esp, 10
'0070d199    8bcc                    mov ecx, esp
'0070d19b    8901                    mov dword ptr [ecx], eax
'0070d19d    8b85d0feffff            mov eax, dword ptr [ebp+fffffed0]
'0070d1a3    894104                  mov dword ptr [ecx+04], eax
'0070d1a6    8b85d4feffff            mov eax, dword ptr [ebp+fffffed4]
'0070d1ac    894108                  mov dword ptr [ecx+08], eax
'0070d1af    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'0070d1b5    89410c                  mov dword ptr [ecx+0c], eax
'0070d1b8    8b45ac                  mov eax, dword ptr [ebp-54]
'0070d1bb    83ec10                  sub esp, 10
'0070d1be    8bcc                    mov ecx, esp
'0070d1c0    8901                    mov dword ptr [ecx], eax
'0070d1c2    8b45b0                  mov eax, dword ptr [ebp-50]
'0070d1c5    894104                  mov dword ptr [ecx+04], eax
'0070d1c8    8b45b4                  mov eax, dword ptr [ebp-4c]
'0070d1cb    894108                  mov dword ptr [ecx+08], eax
'0070d1ce    8b45b8                  mov eax, dword ptr [ebp-48]
'0070d1d1    89410c                  mov dword ptr [ecx+0c], eax
'0070d1d4    8b4de4                  mov ecx, dword ptr [ebp-1c]
'0070d1d7    68209e4300              push 00439e20
'0070d1dc    51                      push ecx
'0070d1dd    ff92f4000000            call dword ptr [edx+000000f4]
'0070d1e3    dbe2                    fnclex
'0070d1e5    85c0                    test ax, ax
'0070d1e7    7d15                    jge 70d1fe
'0070d1e9    8b55e4                  mov edx, dword ptr [ebp-1c]
'0070d1ec    68f4000000              push 000000f4
'0070d1f1    6830314300              push 00433130
'0070d1f6    52                      push edx
'0070d1f7    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070d1f8    ff1580104000            call dword ptr [00401080]
'0070d1fe    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'0070d201    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_43)
'0070d207    8d4dac                  lea ecx, dword ptr [ebp-54]

' *** Reference to "__vbaFreeVar"
'0070d20a    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_50)
'0070d210    8b45e4                  mov eax, dword ptr [ebp-1c]
'0070d213    8b08                    mov ecx, dword ptr [eax]
'0070d215    50                      push eax
'0070d216    ff91d0000000            call dword ptr [ecx+000000d0]
'0070d21c    dbe2                    fnclex
'0070d21e    85c0                    test ax, ax
'0070d220    7d15                    jge 70d237
'0070d222    8b55e4                  mov edx, dword ptr [ebp-1c]
'0070d225    68d0000000              push 000000d0
'0070d22a    6830314300              push 00433130
'0070d22f    52                      push edx
'0070d230    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070d231    ff1580104000            call dword ptr [00401080]
'0070d237    83ec10                  sub esp, 10
'0070d23a    8bd4                    mov edx, esp
'0070d23c    33c0                    xor eax, eax
    var_num1 = Empty
'0070d23e    b903000000              mov ecx, 00000003
'0070d243    890a                    mov dword ptr [edx], ecx
'0070d245    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'0070d24b    894a04                  mov dword ptr [edx+04], ecx
'0070d24e    894208                  mov dword ptr [edx+08], eax
'0070d251    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'0070d257    89420c                  mov dword ptr [edx+0c], eax
'0070d25a    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'0070d260    83ec10                  sub esp, 10
'0070d263    8bcc                    mov ecx, esp
'0070d265    b802000000              mov eax, 00000002
'0070d26a    8901                    mov dword ptr [ecx], eax
'0070d26c    895104                  mov dword ptr [ecx+04], edx
'0070d26f    8b9518ffffff            mov edx, dword ptr [ebp+ffffff18]
'0070d275    66899d14ffffff          mov word ptr [ebp+ffffff14], bx
'0070d27c    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'0070d282    894108                  mov dword ptr [ecx+08], eax
'0070d285    8b85f0feffff            mov eax, dword ptr [ebp+fffffef0]
'0070d28b    89510c                  mov dword ptr [ecx+0c], edx
'0070d28e    83ec10                  sub esp, 10
'0070d291    8bcc                    mov ecx, esp
'0070d293    c785ecfeffff02000000    mov dword ptr [ebp+fffffeec], 00000002
'0070d29d    8b95ecfeffff            mov edx, dword ptr [ebp+fffffeec]
'0070d2a3    8911                    mov dword ptr [ecx], edx
'0070d2a5    8b95f8feffff            mov edx, dword ptr [ebp+fffffef8]
'0070d2ab    894104                  mov dword ptr [ecx+04], eax
'0070d2ae    6a03                    push 03
'0070d2b0    b805000000              mov eax, 00000005
'0070d2b5    894108                  mov dword ptr [ecx+08], eax
'0070d2b8    8b06                    mov eax, dword ptr [esi]
'0070d2ba    689e000000              push 0000009e
'0070d2bf    56                      push esi
'0070d2c0    c785d4feffffcc134300    mov dword ptr [ebp+fffffed4], 004313cc
'0070d2ca    c785ccfeffff08800000    mov dword ptr [ebp+fffffecc], 00008008
'0070d2d4    89510c                  mov dword ptr [ecx+0c], edx
'0070d2d7    ff9020030000            call dword ptr [eax+00000320]
'0070d2dd    50                      push eax
'0070d2de    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070d2e1    51                      push ecx
'0070d2e2    ffd7                    call edi
    Set var_43 = Nothing
'0070d2e4    50                      push eax
'0070d2e5    8d55ac                  lea edx, dword ptr [ebp-54]
'0070d2e8    52                      push edx

' *** Reference to "__vbaLateIdCallLd"
'0070d2e9    ff157c114000            call dword ptr [0040117c]
    var_50 = var_43.UNK_-256 - 12_158
'0070d2ef    83c440                  add esp, 40
'0070d2f2    50                      push eax
'0070d2f3    8d85ccfeffff            lea eax, dword ptr [ebp+fffffecc]
'0070d2f9    50                      push eax

' *** Reference to "__vbaVarTstNe"
'0070d2fa    ff1574124000            call dword ptr [00401274]
'0070d300    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070d303    668985e0fdffff          mov word ptr [ebp+fffffde0], ax

' *** Reference to "__vbaFreeObj"
'0070d30a    ff1508134000            call dword ptr [00401308]
    '#Cleanup(var_43)
'0070d310    8d4dac                  lea ecx, dword ptr [ebp-54]

' *** Reference to "__vbaFreeVar"
'0070d313    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_50)
'0070d319    33c0                    xor eax, eax
    var_num1 = Empty
'0070d31b    663985e0fdffff          cmp word ptr [ebp+fffffde0], ax
'0070d322    0f840c090000            je 70dc34
    
    If (    ((var_50) <> (vbNullChar))) Then
'0070d328    83ec10                  sub esp, 10
'0070d32b    ba02000000              mov edx, 00000002
'0070d330    89953cfeffff            mov dword ptr [ebp+fffffe3c], edx
'0070d336    89958cfeffff            mov dword ptr [ebp+fffffe8c], edx
'0070d33c    8bd4                    mov edx, esp
'0070d33e    b903000000              mov ecx, 00000003
'0070d343    890a                    mov dword ptr [edx], ecx
'0070d345    898dccfeffff            mov dword ptr [ebp+fffffecc], ecx
'0070d34b    8b8dd0feffff            mov ecx, dword ptr [ebp+fffffed0]
'0070d351    894a04                  mov dword ptr [edx+04], ecx
'0070d354    894208                  mov dword ptr [edx+08], eax
'0070d357    898564feffff            mov dword ptr [ebp+fffffe64], eax
'0070d35d    8985d4feffff            mov dword ptr [ebp+fffffed4], eax
'0070d363    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'0070d369    89420c                  mov dword ptr [edx+0c], eax
'0070d36c    8b95b0feffff            mov edx, dword ptr [ebp+fffffeb0]
'0070d372    83ec10                  sub esp, 10
'0070d375    8bcc                    mov ecx, esp
'0070d377    b802000000              mov eax, 00000002
'0070d37c    8901                    mov dword ptr [ecx], eax
'0070d37e    895104                  mov dword ptr [ecx+04], edx
'0070d381    8b95b8feffff            mov edx, dword ptr [ebp+fffffeb8]
'0070d387    66899db4feffff          mov word ptr [ebp+fffffeb4], bx
'0070d38e    8b85b4feffff            mov eax, dword ptr [ebp+fffffeb4]
'0070d394    894108                  mov dword ptr [ecx+08], eax
'0070d397    8b8590feffff            mov eax, dword ptr [ebp+fffffe90]
'0070d39d    89510c                  mov dword ptr [ecx+0c], edx
'0070d3a0    8b958cfeffff            mov edx, dword ptr [ebp+fffffe8c]
'0070d3a6    83ec10                  sub esp, 10
'0070d3a9    8bcc                    mov ecx, esp
'0070d3ab    8911                    mov dword ptr [ecx], edx
'0070d3ad    8b9598feffff            mov edx, dword ptr [ebp+fffffe98]
'0070d3b3    894104                  mov dword ptr [ecx+04], eax
'0070d3b6    6a03                    push 03
'0070d3b8    b805000000              mov eax, 00000005
'0070d3bd    894108                  mov dword ptr [ecx+08], eax
'0070d3c0    8b06                    mov eax, dword ptr [esi]
'0070d3c2    689e000000              push 0000009e
'0070d3c7    56                      push esi
'0070d3c8    c7855cfeffff03000000    mov dword ptr [ebp+fffffe5c], 00000003
'0070d3d2    66899d44feffff          mov word ptr [ebp+fffffe44], bx
'0070d3d9    c78534ffffff12000000    mov dword ptr [ebp+ffffff34], 00000012
'0070d3e3    66899d14ffffff          mov word ptr [ebp+ffffff14], bx
'0070d3ea    89510c                  mov dword ptr [ecx+0c], edx
'0070d3ed    ff9020030000            call dword ptr [eax+00000320]
'0070d3f3    50                      push eax
'0070d3f4    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0070d3f7    51                      push ecx
'0070d3f8    ffd7                    call edi
    Set var_46 = Nothing
'0070d3fa    50                      push eax
'0070d3fb    8d55ac                  lea edx, dword ptr [ebp-54]
'0070d3fe    52                      push edx

' *** Reference to "__vbaLateIdCallLd"
'0070d3ff    ff157c114000            call dword ptr [0040117c]
    var_50 = var_46.UNK_-256 - 12_158
'0070d405    83c440                  add esp, 40
'0070d408    50                      push eax
'0070d409    8d45e0                  lea eax, dword ptr [ebp-20]
'0070d40c    50                      push eax

' *** Reference to "__vbaStrVarVal"
'0070d40d    ff15fc114000            call dword ptr [004011fc]
'0070d413    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'0070d414    ff1510134000            call dword ptr [00401310]
'0070d41a    dd9d74feffff            fstp qword ptr [ebp+fffffe74]
    'var_304 = (00)
'0070d420    8b9530ffffff            mov edx, dword ptr [ebp+ffffff30]
'0070d426    83ec10                  sub esp, 10
'0070d429    8bcc                    mov ecx, esp
'0070d42b    b803000000              mov eax, 00000003
'0070d430    8901                    mov dword ptr [ecx], eax
'0070d432    8b8534ffffff            mov eax, dword ptr [ebp+ffffff34]
'0070d438    895104                  mov dword ptr [ecx+04], edx
'0070d43b    8b9538ffffff            mov edx, dword ptr [ebp+ffffff38]
'0070d441    894108                  mov dword ptr [ecx+08], eax
'0070d444    89510c                  mov dword ptr [ecx+0c], edx
'0070d447    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'0070d44d    83ec10                  sub esp, 10
'0070d450    b802000000              mov eax, 00000002
'0070d455    8bcc                    mov ecx, esp
'0070d457    8901                    mov dword ptr [ecx], eax
'0070d459    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'0070d45f    895104                  mov dword ptr [ecx+04], edx
'0070d462    8b9518ffffff            mov edx, dword ptr [ebp+ffffff18]
'0070d468    c7856cfeffff05000000    mov dword ptr [ebp+fffffe6c], 00000005
'0070d472    894108                  mov dword ptr [ecx+08], eax
'0070d475    89510c                  mov dword ptr [ecx+0c], edx
'0070d478    8b95f0feffff            mov edx, dword ptr [ebp+fffffef0]
'0070d47e    83ec10                  sub esp, 10
'0070d481    8bcc                    mov ecx, esp
'0070d483    b802000000              mov eax, 00000002
'0070d488    8901                    mov dword ptr [ecx], eax
'0070d48a    895104                  mov dword ptr [ecx+04], edx
'0070d48d    b803000000              mov eax, 00000003
'0070d492    894108                  mov dword ptr [ecx+08], eax
'0070d495    8b85f8feffff            mov eax, dword ptr [ebp+fffffef8]
'0070d49b    6a03                    push 03
'0070d49d    689e000000              push 0000009e
'0070d4a2    89410c                  mov dword ptr [ecx+0c], eax
'0070d4a5    8b0e                    mov ecx, dword ptr [esi]
'0070d4a7    56                      push esi
'0070d4a8    ff9120030000            call dword ptr [ecx+00000320]
'0070d4ae    50                      push eax
'0070d4af    8d55cc                  lea edx, dword ptr [ebp-34]
'0070d4b2    52                      push edx
'0070d4b3    ffd7                    call edi
    Set var_43 = Nothing
'0070d4b5    50                      push eax
'0070d4b6    8d459c                  lea eax, dword ptr [ebp-64]
'0070d4b9    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'0070d4ba    ff157c114000            call dword ptr [0040117c]
    var_51 = var_43.UNK_-256 - 12_158
'0070d4c0    83c440                  add esp, 40
'0070d4c3    50                      push eax
'0070d4c4    8d8d6cfeffff            lea ecx, dword ptr [ebp+fffffe6c]
'0070d4ca    51                      push ecx
'0070d4cb    8d558c                  lea edx, dword ptr [ebp-74]
'0070d4ce    52                      push edx

' *** Reference to "__vbaVarAdd"
'0070d4cf    ff158c124000            call dword ptr [0040128c]
'0070d4d5    50                      push eax

' *** Reference to "__vbaStrErrVarCopy"
'0070d4d6    ff1554104000            call dword ptr [00401054]
'0070d4dc    8b955cfeffff            mov edx, dword ptr [ebp+fffffe5c]
'0070d4e2    83ec10                  sub esp, 10
'0070d4e5    8bcc                    mov ecx, esp
'0070d4e7    8911                    mov dword ptr [ecx], edx
'0070d4e9    8b9560feffff            mov edx, dword ptr [ebp+fffffe60]
'0070d4ef    895104                  mov dword ptr [ecx+04], edx
'0070d4f2    8b9564feffff            mov edx, dword ptr [ebp+fffffe64]
'0070d4f8    895108                  mov dword ptr [ecx+08], edx
'0070d4fb    8b9568feffff            mov edx, dword ptr [ebp+fffffe68]
'0070d501    89510c                  mov dword ptr [ecx+0c], edx
'0070d504    8b953cfeffff            mov edx, dword ptr [ebp+fffffe3c]
'0070d50a    83ec10                  sub esp, 10
'0070d50d    8bcc                    mov ecx, esp
'0070d50f    8911                    mov dword ptr [ecx], edx
'0070d511    8b9540feffff            mov edx, dword ptr [ebp+fffffe40]
'0070d517    895104                  mov dword ptr [ecx+04], edx
'0070d51a    8b9544feffff            mov edx, dword ptr [ebp+fffffe44]
'0070d520    895108                  mov dword ptr [ecx+08], edx
'0070d523    8b9548feffff            mov edx, dword ptr [ebp+fffffe48]
'0070d529    89510c                  mov dword ptr [ecx+0c], edx
'0070d52c    83ec10                  sub esp, 10
'0070d52f    8bd4                    mov edx, esp
'0070d531    b902000000              mov ecx, 00000002
'0070d536    890a                    mov dword ptr [edx], ecx
'0070d538    8b8d20feffff            mov ecx, dword ptr [ebp+fffffe20]
'0070d53e    894a04                  mov dword ptr [edx+04], ecx
'0070d541    b903000000              mov ecx, 00000003
'0070d546    894a08                  mov dword ptr [edx+08], ecx
'0070d549    8b8d28feffff            mov ecx, dword ptr [ebp+fffffe28]
'0070d54f    894a0c                  mov dword ptr [edx+0c], ecx
'0070d552    83ec10                  sub esp, 10
'0070d555    8bd4                    mov edx, esp
'0070d557    c7857cffffff08000000    mov dword ptr [ebp+ffffff7c], 00000008
'0070d561    8b8d7cffffff            mov ecx, dword ptr [ebp+ffffff7c]
'0070d567    890a                    mov dword ptr [edx], ecx
'0070d569    8b4d80                  mov ecx, dword ptr [ebp-80]
'0070d56c    894a04                  mov dword ptr [edx+04], ecx
'0070d56f    8b0e                    mov ecx, dword ptr [esi]
'0070d571    6a03                    push 03
'0070d573    894208                  mov dword ptr [edx+08], eax
'0070d576    894584                  mov dword ptr [ebp-7c], eax
'0070d579    8b4588                  mov eax, dword ptr [ebp-78]
'0070d57c    689e000000              push 0000009e
'0070d581    56                      push esi
'0070d582    89420c                  mov dword ptr [edx+0c], eax
'0070d585    ff9120030000            call dword ptr [ecx+00000320]
'0070d58b    50                      push eax
'0070d58c    8d55c4                  lea edx, dword ptr [ebp-3c]
'0070d58f    52                      push edx
'0070d590    ffd7                    call edi
    Set var_9 = Nothing
'0070d592    50                      push eax

' *** Reference to "__vbaLateIdCallSt"
'0070d593    ff159c114000            call dword ptr [0040119c]
'0070d599    83c44c                  add esp, 4c
'0070d59c    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaFreeStr"
'0070d59f    ff150c134000            call dword ptr [0040130c]
    '#Cleanup(var_3)
'0070d5a5    8d45c4                  lea eax, dword ptr [ebp-3c]
'0070d5a8    50                      push eax
'0070d5a9    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0070d5ac    51                      push ecx
'0070d5ad    8d55cc                  lea edx, dword ptr [ebp-34]
'0070d5b0    52                      push edx
'0070d5b1    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'0070d5b3    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_43, var_46, var_9)
'0070d5b9    8d857cffffff            lea eax, dword ptr [ebp+ffffff7c]
'0070d5bf    50                      push eax
'0070d5c0    8d4d8c                  lea ecx, dword ptr [ebp-74]
'0070d5c3    51                      push ecx
'0070d5c4    8d558c                  lea edx, dword ptr [ebp-74]
'0070d5c7    52                      push edx
'0070d5c8    8d459c                  lea eax, dword ptr [ebp-64]
'0070d5cb    50                      push eax
'0070d5cc    8d4dac                  lea ecx, dword ptr [ebp-54]
'0070d5cf    51                      push ecx
'0070d5d0    6a05                    push 05

' *** Reference to "__vbaFreeVarList"
'0070d5d2    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_50, var_51, var_52, var_52, var_59)
'0070d5d8    83c418                  add esp, 18
'0070d5db    66899d14ffffff          mov word ptr [ebp+ffffff14], bx
'0070d5e2    8bdc                    mov ebx, esp
'0070d5e4    33c0                    xor eax, eax
'0070d5e6    b903000000              mov ecx, 00000003
'0070d5eb    890b                    mov dword ptr [ebx], ecx
'0070d5ed    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'0070d5f3    894b04                  mov dword ptr [ebx+04], ecx
'0070d5f6    894308                  mov dword ptr [ebx+08], eax
'0070d5f9    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'0070d5ff    89430c                  mov dword ptr [ebx+0c], eax
'0070d602    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'0070d608    83ec10                  sub esp, 10
'0070d60b    8bcc                    mov ecx, esp
'0070d60d    ba02000000              mov edx, 00000002
'0070d612    8911                    mov dword ptr [ecx], edx
'0070d614    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'0070d61a    895104                  mov dword ptr [ecx+04], edx
'0070d61d    8b9518ffffff            mov edx, dword ptr [ebp+ffffff18]
'0070d623    894108                  mov dword ptr [ecx+08], eax
'0070d626    89510c                  mov dword ptr [ecx+0c], edx
'0070d629    8b95f0feffff            mov edx, dword ptr [ebp+fffffef0]
'0070d62f    83ec10                  sub esp, 10
'0070d632    8bcc                    mov ecx, esp
'0070d634    b802000000              mov eax, 00000002
'0070d639    8901                    mov dword ptr [ecx], eax
'0070d63b    895104                  mov dword ptr [ecx+04], edx
'0070d63e    b805000000              mov eax, 00000005
'0070d643    894108                  mov dword ptr [ecx+08], eax
'0070d646    8b85f8feffff            mov eax, dword ptr [ebp+fffffef8]
'0070d64c    6a03                    push 03
'0070d64e    689e000000              push 0000009e
'0070d653    89410c                  mov dword ptr [ecx+0c], eax
'0070d656    8b0e                    mov ecx, dword ptr [esi]
'0070d658    56                      push esi
'0070d659    ff9120030000            call dword ptr [ecx+00000320]
'0070d65f    50                      push eax
'0070d660    8d55c8                  lea edx, dword ptr [ebp-38]
'0070d663    52                      push edx
'0070d664    ffd7                    call edi
    Set var_46 = Nothing
'0070d666    50                      push eax
'0070d667    8d45ac                  lea eax, dword ptr [ebp-54]
'0070d66a    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'0070d66b    ff157c114000            call dword ptr [0040117c]
    var_50 = var_46.UNK_-256 - 12_158
'0070d671    83c440                  add esp, 40
'0070d674    50                      push eax
'0070d675    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0070d678    51                      push ecx

' *** Reference to "__vbaStrVarVal"
'0070d679    ff15fc114000            call dword ptr [004011fc]
'0070d67f    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'0070d680    ff1510134000            call dword ptr [00401310]
'0070d686    dd9df4fdffff            fstp qword ptr [ebp+fffffdf4]
    'var_494 = (00)
'0070d68c    8b16                    mov edx, dword ptr [esi]
'0070d68e    56                      push esi
'0070d68f    ff92fc020000            call dword ptr [edx+000002fc]
'0070d695    50                      push eax
'0070d696    8d45c4                  lea eax, dword ptr [ebp-3c]
'0070d699    50                      push eax
'0070d69a    ffd7                    call edi
    Set var_9 = 
'0070d69c    8d55d8                  lea edx, dword ptr [ebp-28]
'0070d69f    8bd8                    mov ebx, eax
'0070d6a1    8b0b                    mov ecx, dword ptr [ebx]
'0070d6a3    52                      push edx
'0070d6a4    53                      push ebx
'0070d6a5    ff91a0000000            call dword ptr [ecx+000000a0]
'0070d6ab    dbe2                    fnclex
'0070d6ad    85c0                    test ax, ax
'0070d6af    7d12                    jge 70d6c3
'0070d6b1    68a0000000              push 000000a0
'0070d6b6    68d00d4300              push 00430dd0
'0070d6bb    53                      push ebx
'0070d6bc    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070d6bd    ff1580104000            call dword ptr [00401080]
'0070d6c3    8b45d8                  mov eax, dword ptr [ebp-28]
'0070d6c6    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'0070d6c7    ff1510134000            call dword ptr [00401310]
'0070d6cd    dd9decfdffff            fstp qword ptr [ebp+fffffdec]
    'var_293 = (00)
'0070d6d3    8b0e                    mov ecx, dword ptr [esi]
'0070d6d5    56                      push esi
'0070d6d6    ff91fc020000            call dword ptr [ecx+000002fc]
'0070d6dc    50                      push eax
'0070d6dd    8d55c0                  lea edx, dword ptr [ebp-40]
'0070d6e0    52                      push edx
'0070d6e1    ffd7                    call edi
    Set var_5 = Nothing
'0070d6e3    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0070d6e6    8bd8                    mov ebx, eax
'0070d6e8    8b03                    mov eax, dword ptr [ebx]
'0070d6ea    51                      push ecx
'0070d6eb    53                      push ebx
'0070d6ec    ff90a0000000            call dword ptr [eax+000000a0]
'0070d6f2    dbe2                    fnclex
'0070d6f4    85c0                    test ax, ax
'0070d6f6    7d12                    jge 70d70a
'0070d6f8    68a0000000              push 000000a0
'0070d6fd    68d00d4300              push 00430dd0
'0070d702    53                      push ebx
'0070d703    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070d704    ff1580104000            call dword ptr [00401080]
'0070d70a    8b55d4                  mov edx, dword ptr [ebp-2c]
'0070d70d    52                      push edx

' *** Reference to "rtcR8ValFromBstr"
'0070d70e    ff1510134000            call dword ptr [00401310]
'0070d714    dd9de4fdffff            fstp qword ptr [ebp+fffffde4]
    'var_495 = (00)
'0070d71a    8b06                    mov eax, dword ptr [esi]
'0070d71c    56                      push esi
'0070d71d    ff9018030000            call dword ptr [eax+00000318]
'0070d723    50                      push eax
'0070d724    8d4dbc                  lea ecx, dword ptr [ebp-44]
'0070d727    51                      push ecx
'0070d728    ffd7                    call edi
    Set var_58 = Nothing
'0070d72a    8b16                    mov edx, dword ptr [esi]
'0070d72c    56                      push esi
'0070d72d    8985c8fdffff            mov dword ptr [ebp+fffffdc8], eax
'0070d733    ff9218030000            call dword ptr [edx+00000318]
'0070d739    50                      push eax
'0070d73a    8d45cc                  lea eax, dword ptr [ebp-34]
'0070d73d    50                      push eax
'0070d73e    ffd7                    call edi
    Set var_43 = Nothing
'0070d740    8d55e0                  lea edx, dword ptr [ebp-20]
'0070d743    8bd8                    mov ebx, eax
'0070d745    8b0b                    mov ecx, dword ptr [ebx]
'0070d747    52                      push edx
'0070d748    53                      push ebx
'0070d749    ff5150                  call dword ptr [ecx+50]
'0070d74c    dbe2                    fnclex
'0070d74e    85c0                    test ax, ax
'0070d750    7d0f                    jge 70d761
'0070d752    6a50                    push 50
'0070d754    683ce04300              push 0043e03c
'0070d759    53                      push ebx
'0070d75a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070d75b    ff1580104000            call dword ptr [00401080]
'0070d761    dd85ecfdffff            fld qword ptr [ebp+fffffdec]
'0070d767    8b85c8fdffff            mov eax, dword ptr [ebp+fffffdc8]
'0070d76d    dc8df4fdffff            fmul qword ptr [ebp+fffffdf4]
'0070d773    8b18                    mov ebx, dword ptr [eax]
'0070d775    dd05286c4000            fld qword ptr [00406c28]
'0070d77b    dca5e4fdffff            fsub qword ptr [ebp+fffffde4]
'0070d781    833d00b0720000          cmp dword ptr [0072b000], 00
'0070d788    7504                    jne 70d78e
    
    If (    vbNullChar = 0) Then
'0070d78a    def9                    fdivp
'0070d78c    eb09                    jmp 70d797
    
End If
'0070d78e    50                      push eax
'0070d78f    b00d                    mov al, 0d

' *** Reference to "_adj_fdiv_r"
'0070d791    e8f49acfff              call 40728a
'0070d796    58                      pop eax
'0070d797    dfe0                    fnstsw ax
'0070d799    a80d                    test al, 0d
'0070d79b    0f855e0b0000            jne 70e2ff

' *** Reference to "__vbaFPInt"
'0070d7a1    ff15fc124000            call dword ptr [004012fc]
'0070d7a7    dd9da4fdffff            fstp qword ptr [ebp+fffffda4]
'var_499 = (00)
'0070d7ad    8b4de0                  mov ecx, dword ptr [ebp-20]
'0070d7b0    51                      push ecx

' *** Reference to "rtcR8ValFromBstr"
'0070d7b1    ff1510134000            call dword ptr [00401310]
'0070d7b7    dc85a4fdffff            fadd qword ptr [ebp+fffffda4]
'0070d7bd    83ec08                  sub esp, 08
'0070d7c0    dfe0                    fnstsw ax
'0070d7c2    a80d                    test al, 0d
'0070d7c4    0f85350b0000            jne 70e2ff
'0070d7ca    dd1c24                  fstp qword ptr [esp]
'var_922 = (00)

' *** Reference to "__vbaStrR8"
'0070d7cd    ff1580114000            call dword ptr [00401180]
'0070d7d3    8bd0                    mov edx, eax
'0070d7d5    8d4dd0                  lea ecx, dword ptr [ebp-30]

' *** Reference to "__vbaStrMove"
'0070d7d8    ff15d0124000            call dword ptr [004012d0]
'0070d7de    8bd3                    mov edx, ebx
'0070d7e0    8b9dc8fdffff            mov ebx, dword ptr [ebp+fffffdc8]
'0070d7e6    50                      push eax
'0070d7e7    53                      push ebx
'0070d7e8    ff5254                  call dword ptr [edx+54]
'0070d7eb    dbe2                    fnclex
'0070d7ed    85c0                    test ax, ax
'0070d7ef    7d0f                    jge 70d800
'0070d7f1    6a54                    push 54
'0070d7f3    683ce04300              push 0043e03c
'0070d7f8    53                      push ebx
'0070d7f9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070d7fa    ff1580104000            call dword ptr [00401080]
'0070d800    8d45d0                  lea eax, dword ptr [ebp-30]
'0070d803    50                      push eax
'0070d804    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0070d807    51                      push ecx
'0070d808    8d55d8                  lea edx, dword ptr [ebp-28]
'0070d80b    52                      push edx
'0070d80c    8d45dc                  lea eax, dword ptr [ebp-24]
'0070d80f    50                      push eax
'0070d810    8d4de0                  lea ecx, dword ptr [ebp-20]
'0070d813    51                      push ecx
'0070d814    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'0070d816    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , 0, 0, 0, -4680)
'0070d81c    8d55bc                  lea edx, dword ptr [ebp-44]
'0070d81f    52                      push edx
'0070d820    8d45c0                  lea eax, dword ptr [ebp-40]
'0070d823    50                      push eax
'0070d824    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0070d827    51                      push ecx
'0070d828    8d55c8                  lea edx, dword ptr [ebp-38]
'0070d82b    52                      push edx
'0070d82c    8d45cc                  lea eax, dword ptr [ebp-34]
'0070d82f    50                      push eax
'0070d830    6a05                    push 05

' *** Reference to "__vbaFreeObjList"
'0070d832    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46, var_9, var_5, var_58)
'0070d838    83c430                  add esp, 30
'0070d83b    8d4dac                  lea ecx, dword ptr [ebp-54]

' *** Reference to "__vbaFreeVar"
'0070d83e    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_50)
'0070d844    8b0e                    mov ecx, dword ptr [esi]
'0070d846    56                      push esi
'0070d847    ff910c030000            call dword ptr [ecx+0000030c]
'0070d84d    50                      push eax
'0070d84e    8d55c0                  lea edx, dword ptr [ebp-40]
'0070d851    52                      push edx
'0070d852    ffd7                    call edi
Set var_5 = 
'0070d854    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0070d85a    8b06                    mov eax, dword ptr [esi]
'0070d85c    56                      push esi
'0070d85d    ff900c030000            call dword ptr [eax+0000030c]
'0070d863    50                      push eax
'0070d864    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0070d867    51                      push ecx
'0070d868    ffd7                    call edi
Set var_46 = Nothing
'0070d86a    8bd8                    mov ebx, eax
'0070d86c    8b13                    mov edx, dword ptr [ebx]
'0070d86e    8d45e0                  lea eax, dword ptr [ebp-20]
'0070d871    50                      push eax
'0070d872    53                      push ebx
'0070d873    ff92a0000000            call dword ptr [edx+000000a0]
'0070d879    dbe2                    fnclex
'0070d87b    85c0                    test ax, ax
'0070d87d    7d12                    jge 70d891
'0070d87f    68a0000000              push 000000a0
'0070d884    68d00d4300              push 00430dd0
'0070d889    53                      push ebx
'0070d88a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070d88b    ff1580104000            call dword ptr [00401080]
'0070d891    8b55e8                  mov edx, dword ptr [ebp-18]
'0070d894    8b45e0                  mov eax, dword ptr [ebp-20]
'0070d897    668995a4feffff          mov word ptr [ebp+fffffea4], dx
'0070d89e    66899514ffffff          mov word ptr [ebp+ffffff14], dx
'0070d8a5    83ec10                  sub esp, 10
'0070d8a8    8945b4                  mov dword ptr [ebp-4c], eax
'0070d8ab    8bd4                    mov edx, esp
'0070d8ad    b903000000              mov ecx, 00000003
'0070d8b2    890a                    mov dword ptr [edx], ecx
'0070d8b4    8b8dc0feffff            mov ecx, dword ptr [ebp+fffffec0]
'0070d8ba    894a04                  mov dword ptr [edx+04], ecx
'0070d8bd    b808000000              mov eax, 00000008
'0070d8c2    8945ac                  mov dword ptr [ebp-54], eax
'0070d8c5    8985ccfeffff            mov dword ptr [ebp+fffffecc], eax
'0070d8cb    33c0                    xor eax, eax
var_num1 = Empty
'0070d8cd    894208                  mov dword ptr [edx+08], eax
'0070d8d0    8b85c8feffff            mov eax, dword ptr [ebp+fffffec8]
'0070d8d6    89420c                  mov dword ptr [edx+0c], eax
'0070d8d9    8b95a0feffff            mov edx, dword ptr [ebp+fffffea0]
'0070d8df    83ec10                  sub esp, 10
'0070d8e2    8bcc                    mov ecx, esp
'0070d8e4    b802000000              mov eax, 00000002
'0070d8e9    8901                    mov dword ptr [ecx], eax
'0070d8eb    8b85a4feffff            mov eax, dword ptr [ebp+fffffea4]
'0070d8f1    895104                  mov dword ptr [ecx+04], edx
'0070d8f4    8b95a8feffff            mov edx, dword ptr [ebp+fffffea8]
'0070d8fa    894108                  mov dword ptr [ecx+08], eax
'0070d8fd    8b8580feffff            mov eax, dword ptr [ebp+fffffe80]
'0070d903    89510c                  mov dword ptr [ecx+0c], edx
'0070d906    83ec10                  sub esp, 10
'0070d909    8bcc                    mov ecx, esp
'0070d90b    c7857cfeffff02000000    mov dword ptr [ebp+fffffe7c], 00000002
'0070d915    8b957cfeffff            mov edx, dword ptr [ebp+fffffe7c]
'0070d91b    8911                    mov dword ptr [ecx], edx
'0070d91d    8b9588feffff            mov edx, dword ptr [ebp+fffffe88]
'0070d923    894104                  mov dword ptr [ecx+04], eax
'0070d926    6a03                    push 03
'0070d928    b805000000              mov eax, 00000005
'0070d92d    894108                  mov dword ptr [ecx+08], eax
'0070d930    8b06                    mov eax, dword ptr [esi]
'0070d932    689e000000              push 0000009e
'0070d937    33db                    xor ebx, ebx
var_num2 = Empty
'0070d939    56                      push esi
'0070d93a    895de0                  mov dword ptr [ebp-20], ebx
'0070d93d    c785d4feffffe42b4400    mov dword ptr [ebp+fffffed4], 00442be4
'0070d947    89510c                  mov dword ptr [ecx+0c], edx
'0070d94a    ff9020030000            call dword ptr [eax+00000320]
'0070d950    50                      push eax
'0070d951    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0070d954    51                      push ecx
'0070d955    ffd7                    call edi
Set var_9 = Nothing
'0070d957    50                      push eax
'0070d958    8d956cffffff            lea edx, dword ptr [ebp+ffffff6c]
'0070d95e    52                      push edx

' *** Reference to "__vbaLateIdCallLd"
'0070d95f    ff157c114000            call dword ptr [0040117c]
var_94 = var_9.UNK_-256 - 12_158
'0070d965    83c440                  add esp, 40
'0070d968    50                      push eax
'0070d969    8d45dc                  lea eax, dword ptr [ebp-24]
'0070d96c    50                      push eax

' *** Reference to "__vbaStrVarVal"
'0070d96d    ff15fc114000            call dword ptr [004011fc]
'0070d973    50                      push eax

' *** Reference to "rtcR8ValFromBstr"
'0070d974    ff1510134000            call dword ptr [00401310]
'0070d97a    dd9d64feffff            fstp qword ptr [ebp+fffffe64]
'var_122 = (00)
'0070d980    8b8dd8fdffff            mov ecx, dword ptr [ebp+fffffdd8]
'0070d986    b808000000              mov eax, 00000008
'0070d98b    89854cfeffff            mov dword ptr [ebp+fffffe4c], eax
'0070d991    89853cfeffff            mov dword ptr [ebp+fffffe3c], eax
'0070d997    8d45ac                  lea eax, dword ptr [ebp-54]
'0070d99a    50                      push eax
'0070d99b    c7855cfeffff05000000    mov dword ptr [ebp+fffffe5c], 00000005
'0070d9a5    c78554feffffa0654500    mov dword ptr [ebp+fffffe54], 004565a0
'0070d9af    c78544feffff44ed4300    mov dword ptr [ebp+fffffe44], 0043ed44
'0070d9b9    8b11                    mov edx, dword ptr [ecx]
'0070d9bb    83ec10                  sub esp, 10
'0070d9be    b803000000              mov eax, 00000003
'0070d9c3    8bcc                    mov ecx, esp
'0070d9c5    8901                    mov dword ptr [ecx], eax
'0070d9c7    8b8530ffffff            mov eax, dword ptr [ebp+ffffff30]
'0070d9cd    894104                  mov dword ptr [ecx+04], eax
'0070d9d0    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'0070d9d6    895908                  mov dword ptr [ecx+08], ebx
'0070d9d9    89410c                  mov dword ptr [ecx+0c], eax
'0070d9dc    83ec10                  sub esp, 10
'0070d9df    8bcc                    mov ecx, esp
'0070d9e1    b802000000              mov eax, 00000002
'0070d9e6    8901                    mov dword ptr [ecx], eax
'0070d9e8    8b8510ffffff            mov eax, dword ptr [ebp+ffffff10]
'0070d9ee    894104                  mov dword ptr [ecx+04], eax
'0070d9f1    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'0070d9f7    894108                  mov dword ptr [ecx+08], eax
'0070d9fa    8b8518ffffff            mov eax, dword ptr [ebp+ffffff18]
'0070da00    89410c                  mov dword ptr [ecx+0c], eax
'0070da03    83ec10                  sub esp, 10
'0070da06    8bcc                    mov ecx, esp
'0070da08    b802000000              mov eax, 00000002
'0070da0d    8901                    mov dword ptr [ecx], eax
'0070da0f    8b85f0feffff            mov eax, dword ptr [ebp+fffffef0]
'0070da15    894104                  mov dword ptr [ecx+04], eax
'0070da18    33c0                    xor eax, eax
var_num1 = Empty
'0070da1a    894108                  mov dword ptr [ecx+08], eax
'0070da1d    8b85f8feffff            mov eax, dword ptr [ebp+fffffef8]
'0070da23    6a03                    push 03
'0070da25    689e000000              push 0000009e
'0070da2a    89410c                  mov dword ptr [ecx+0c], eax
'0070da2d    8b0e                    mov ecx, dword ptr [esi]
'0070da2f    56                      push esi
'0070da30    89959cfdffff            mov dword ptr [ebp+fffffd9c], edx
'0070da36    ff9120030000            call dword ptr [ecx+00000320]
'0070da3c    50                      push eax
'0070da3d    8d55cc                  lea edx, dword ptr [ebp-34]
'0070da40    52                      push edx
'0070da41    ffd7                    call edi
Set var_43 = Nothing
'0070da43    50                      push eax
'0070da44    8d459c                  lea eax, dword ptr [ebp-64]
'0070da47    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'0070da48    ff157c114000            call dword ptr [0040117c]
var_51 = var_43.UNK_-256 - 12_158
'0070da4e    83c440                  add esp, 40
'0070da51    50                      push eax
'0070da52    8d4d8c                  lea ecx, dword ptr [ebp-74]
'0070da55    51                      push ecx

' *** Reference to "__vbaVarCat"
'0070da56    ff1508124000            call dword ptr [00401208]
'0070da5c    50                      push eax
'0070da5d    8d95ccfeffff            lea edx, dword ptr [ebp+fffffecc]
'0070da63    52                      push edx
'0070da64    8d857cffffff            lea eax, dword ptr [ebp+ffffff7c]
'0070da6a    50                      push eax

' *** Reference to "__vbaVarCat"
'0070da6b    ff1508124000            call dword ptr [00401208]
'0070da71    50                      push eax
'0070da72    8d8d5cfeffff            lea ecx, dword ptr [ebp+fffffe5c]
'0070da78    51                      push ecx
'0070da79    8d955cffffff            lea edx, dword ptr [ebp+ffffff5c]
'0070da7f    52                      push edx

' *** Reference to "__vbaVarCat"
'0070da80    ff1508124000            call dword ptr [00401208]
'0070da86    50                      push eax
'0070da87    8d854cfeffff            lea eax, dword ptr [ebp+fffffe4c]
'0070da8d    50                      push eax
'0070da8e    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'0070da94    51                      push ecx

' *** Reference to "__vbaVarCat"
'0070da95    ff1508124000            call dword ptr [00401208]
'0070da9b    50                      push eax
'0070da9c    8d953cfeffff            lea edx, dword ptr [ebp+fffffe3c]
'0070daa2    52                      push edx
'0070daa3    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'0070daa9    50                      push eax

' *** Reference to "__vbaVarCat"
'0070daaa    ff1508124000            call dword ptr [00401208]
'0070dab0    50                      push eax
'0070dab1    8d4dd8                  lea ecx, dword ptr [ebp-28]
'0070dab4    51                      push ecx

' *** Reference to "__vbaStrVarVal"
'0070dab5    ff15fc114000            call dword ptr [004011fc]
'0070dabb    8b9dd8fdffff            mov ebx, dword ptr [ebp+fffffdd8]
'0070dac1    8b959cfdffff            mov edx, dword ptr [ebp+fffffd9c]
'0070dac7    50                      push eax
'0070dac8    53                      push ebx
'0070dac9    ff92a4000000            call dword ptr [edx+000000a4]
'0070dacf    dbe2                    fnclex
'0070dad1    85c0                    test ax, ax
'0070dad3    7d12                    jge 70dae7
'0070dad5    68a4000000              push 000000a4
'0070dada    68d00d4300              push 00430dd0
'0070dadf    53                      push ebx
'0070dae0    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070dae1    ff1580104000            call dword ptr [00401080]
'0070dae7    8d45d8                  lea eax, dword ptr [ebp-28]
'0070daea    50                      push eax
'0070daeb    8d4ddc                  lea ecx, dword ptr [ebp-24]
'0070daee    51                      push ecx
'0070daef    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'0070daf1    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, 0)
'0070daf7    8d55c0                  lea edx, dword ptr [ebp-40]
'0070dafa    52                      push edx
'0070dafb    8d45c4                  lea eax, dword ptr [ebp-3c]
'0070dafe    50                      push eax
'0070daff    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0070db02    51                      push ecx
'0070db03    8d55cc                  lea edx, dword ptr [ebp-34]
'0070db06    52                      push edx
'0070db07    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'0070db09    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46, var_9, var_5)
'0070db0f    8d853cffffff            lea eax, dword ptr [ebp+ffffff3c]
'0070db15    50                      push eax
'0070db16    8d8d4cffffff            lea ecx, dword ptr [ebp+ffffff4c]
'0070db1c    51                      push ecx
'0070db1d    8d955cffffff            lea edx, dword ptr [ebp+ffffff5c]
'0070db23    52                      push edx
'0070db24    8d857cffffff            lea eax, dword ptr [ebp+ffffff7c]
'0070db2a    50                      push eax
'0070db2b    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'0070db31    51                      push ecx
'0070db32    8d558c                  lea edx, dword ptr [ebp-74]
'0070db35    52                      push edx
'0070db36    8d459c                  lea eax, dword ptr [ebp-64]
'0070db39    50                      push eax
'0070db3a    8d4dac                  lea ecx, dword ptr [ebp-54]
'0070db3d    51                      push ecx
'0070db3e    6a08                    push 08

' *** Reference to "__vbaFreeVarList"
'0070db40    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_94, var_59, var_88, var_89, var_90)
'0070db46    8b5de8                  mov ebx, dword ptr [ebp-18]
'0070db49    83c444                  add esp, 44
'0070db4c    83ec10                  sub esp, 10
'0070db4f    33c0                    xor eax, eax
'0070db51    8bd4                    mov edx, esp
'0070db53    b903000000              mov ecx, 00000003
'0070db58    890a                    mov dword ptr [edx], ecx
'0070db5a    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'0070db60    894a04                  mov dword ptr [edx+04], ecx
'0070db63    894208                  mov dword ptr [edx+08], eax
'0070db66    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'0070db6c    89420c                  mov dword ptr [edx+0c], eax
'0070db6f    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'0070db75    83ec10                  sub esp, 10
'0070db78    8bcc                    mov ecx, esp
'0070db7a    b802000000              mov eax, 00000002
'0070db7f    8901                    mov dword ptr [ecx], eax
'0070db81    895104                  mov dword ptr [ecx+04], edx
'0070db84    8b9518ffffff            mov edx, dword ptr [ebp+ffffff18]
'0070db8a    66899d14ffffff          mov word ptr [ebp+ffffff14], bx
'0070db91    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'0070db97    894108                  mov dword ptr [ecx+08], eax
'0070db9a    8b85f0feffff            mov eax, dword ptr [ebp+fffffef0]
'0070dba0    89510c                  mov dword ptr [ecx+0c], edx
'0070dba3    83ec10                  sub esp, 10
'0070dba6    8bcc                    mov ecx, esp
'0070dba8    c785ecfeffff02000000    mov dword ptr [ebp+fffffeec], 00000002
'0070dbb2    8b95ecfeffff            mov edx, dword ptr [ebp+fffffeec]
'0070dbb8    8911                    mov dword ptr [ecx], edx
'0070dbba    8b95f8feffff            mov edx, dword ptr [ebp+fffffef8]
'0070dbc0    894104                  mov dword ptr [ecx+04], eax
'0070dbc3    b805000000              mov eax, 00000005
'0070dbc8    894108                  mov dword ptr [ecx+08], eax
'0070dbcb    89510c                  mov dword ptr [ecx+0c], edx
'0070dbce    8b95d0feffff            mov edx, dword ptr [ebp+fffffed0]
'0070dbd4    83ec10                  sub esp, 10
'0070dbd7    c785ccfeffff08000000    mov dword ptr [ebp+fffffecc], 00000008
'0070dbe1    8b8dccfeffff            mov ecx, dword ptr [ebp+fffffecc]
'0070dbe7    8bc4                    mov eax, esp
'0070dbe9    8908                    mov dword ptr [eax], ecx
'0070dbeb    895004                  mov dword ptr [eax+04], edx
'0070dbee    8b95d8feffff            mov edx, dword ptr [ebp+fffffed8]
'0070dbf4    c785d4feffffcc134300    mov dword ptr [ebp+fffffed4], 004313cc
'0070dbfe    8b8dd4feffff            mov ecx, dword ptr [ebp+fffffed4]
'0070dc04    894808                  mov dword ptr [eax+08], ecx
'0070dc07    89500c                  mov dword ptr [eax+0c], edx
'0070dc0a    6a03                    push 03
'0070dc0c    8b06                    mov eax, dword ptr [esi]
'0070dc0e    689e000000              push 0000009e
'0070dc13    56                      push esi
'0070dc14    ff9020030000            call dword ptr [eax+00000320]
'0070dc1a    50                      push eax
'0070dc1b    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070dc1e    51                      push ecx
'0070dc1f    ffd7                    call edi
Set var_43 = Nothing
'0070dc21    50                      push eax

' *** Reference to "__vbaLateIdCallSt"
'0070dc22    ff159c114000            call dword ptr [0040119c]
'0070dc28    83c44c                  add esp, 4c
'0070dc2b    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'0070dc2e    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'ERROR: Two many next close:
End If
'0070dc34    8b55e4                  mov edx, dword ptr [ebp-1c]
'0070dc37    83ec10                  sub esp, 10
'0070dc3a    66899d14ffffff          mov word ptr [ebp+ffffff14], bx
'0070dc41    c785d4feffff84af4300    mov dword ptr [ebp+fffffed4], 0043af84
'0070dc4b    c785ccfeffff08000000    mov dword ptr [ebp+fffffecc], 00000008
'0070dc55    8b1a                    mov ebx, dword ptr [edx]
'0070dc57    8bd4                    mov edx, esp
'0070dc59    b903000000              mov ecx, 00000003
'0070dc5e    890a                    mov dword ptr [edx], ecx
'0070dc60    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'0070dc66    894a04                  mov dword ptr [edx+04], ecx
'0070dc69    b812000000              mov eax, 00000012
'0070dc6e    894208                  mov dword ptr [edx+08], eax
'0070dc71    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'0070dc77    89420c                  mov dword ptr [edx+0c], eax
'0070dc7a    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'0070dc80    83ec10                  sub esp, 10
'0070dc83    8bcc                    mov ecx, esp
'0070dc85    b802000000              mov eax, 00000002
'0070dc8a    8901                    mov dword ptr [ecx], eax
'0070dc8c    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'0070dc92    895104                  mov dword ptr [ecx+04], edx
'0070dc95    8b9518ffffff            mov edx, dword ptr [ebp+ffffff18]
'0070dc9b    894108                  mov dword ptr [ecx+08], eax
'0070dc9e    8b85f0feffff            mov eax, dword ptr [ebp+fffffef0]
'0070dca4    89510c                  mov dword ptr [ecx+0c], edx
'0070dca7    83ec10                  sub esp, 10
'0070dcaa    8bcc                    mov ecx, esp
'0070dcac    c785ecfeffff02000000    mov dword ptr [ebp+fffffeec], 00000002
'0070dcb6    8b95ecfeffff            mov edx, dword ptr [ebp+fffffeec]
'0070dcbc    8911                    mov dword ptr [ecx], edx
'0070dcbe    8b95f8feffff            mov edx, dword ptr [ebp+fffffef8]
'0070dcc4    894104                  mov dword ptr [ecx+04], eax
'0070dcc7    b803000000              mov eax, 00000003
'0070dccc    50                      push eax
'0070dccd    894108                  mov dword ptr [ecx+08], eax
'0070dcd0    8b06                    mov eax, dword ptr [esi]
'0070dcd2    689e000000              push 0000009e
'0070dcd7    56                      push esi
'0070dcd8    89510c                  mov dword ptr [ecx+0c], edx
'0070dcdb    ff9020030000            call dword ptr [eax+00000320]
'0070dce1    50                      push eax
'0070dce2    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070dce5    51                      push ecx
'0070dce6    ffd7                    call edi
Set var_43 = Nothing
'0070dce8    50                      push eax
'0070dce9    8d55ac                  lea edx, dword ptr [ebp-54]
'0070dcec    52                      push edx

' *** Reference to "__vbaLateIdCallLd"
'0070dced    ff157c114000            call dword ptr [0040117c]
var_50 = var_43.UNK_-256 - 12_158
'0070dcf3    8b10                    mov edx, dword ptr [eax]
'0070dcf5    83c430                  add esp, 30
'0070dcf8    8bcc                    mov ecx, esp
'0070dcfa    8911                    mov dword ptr [ecx], edx
'0070dcfc    8b5004                  mov edx, dword ptr [eax+04]
'0070dcff    895104                  mov dword ptr [ecx+04], edx
'0070dd02    8b5008                  mov edx, dword ptr [eax+08]
'0070dd05    8b400c                  mov eax, dword ptr [eax+0c]
'0070dd08    895108                  mov dword ptr [ecx+08], edx
'0070dd0b    8b95ccfeffff            mov edx, dword ptr [ebp+fffffecc]
'0070dd11    89410c                  mov dword ptr [ecx+0c], eax
'0070dd14    8b85d0feffff            mov eax, dword ptr [ebp+fffffed0]
'0070dd1a    83ec10                  sub esp, 10
'0070dd1d    8bcc                    mov ecx, esp
'0070dd1f    8911                    mov dword ptr [ecx], edx
'0070dd21    8b95d4feffff            mov edx, dword ptr [ebp+fffffed4]
'0070dd27    894104                  mov dword ptr [ecx+04], eax
'0070dd2a    8b85d8feffff            mov eax, dword ptr [ebp+fffffed8]
'0070dd30    895108                  mov dword ptr [ecx+08], edx
'0070dd33    89410c                  mov dword ptr [ecx+0c], eax
'0070dd36    8b4de4                  mov ecx, dword ptr [ebp-1c]
'0070dd39    51                      push ecx
'0070dd3a    ff9328010000            call dword ptr [ebx+00000128]
'0070dd40    dbe2                    fnclex
'0070dd42    85c0                    test ax, ax
'0070dd44    7d15                    jge 70dd5b
'0070dd46    8b55e4                  mov edx, dword ptr [ebp-1c]
'0070dd49    6828010000              push 00000128
'0070dd4e    6830314300              push 00433130
'0070dd53    52                      push edx
'0070dd54    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070dd55    ff1580104000            call dword ptr [00401080]
'0070dd5b    8d4dcc                  lea ecx, dword ptr [ebp-34]

' *** Reference to "__vbaFreeObj"
'0070dd5e    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_43)
'0070dd64    8d4dac                  lea ecx, dword ptr [ebp-54]

' *** Reference to "__vbaFreeVar"
'0070dd67    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_50)
'0070dd6d    8b5de8                  mov ebx, dword ptr [ebp-18]
'0070dd70    83ec10                  sub esp, 10
'0070dd73    8bd4                    mov edx, esp
'0070dd75    b903000000              mov ecx, 00000003
'0070dd7a    890a                    mov dword ptr [edx], ecx
'0070dd7c    898dccfeffff            mov dword ptr [ebp+fffffecc], ecx
'0070dd82    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'0070dd88    894a04                  mov dword ptr [edx+04], ecx
'0070dd8b    b812000000              mov eax, 00000012
'0070dd90    894208                  mov dword ptr [edx+08], eax
'0070dd93    8985d4feffff            mov dword ptr [ebp+fffffed4], eax
'0070dd99    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'0070dd9f    89420c                  mov dword ptr [edx+0c], eax
'0070dda2    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'0070dda8    83ec10                  sub esp, 10
'0070ddab    8bcc                    mov ecx, esp
'0070ddad    b802000000              mov eax, 00000002
'0070ddb2    8901                    mov dword ptr [ecx], eax
'0070ddb4    895104                  mov dword ptr [ecx+04], edx
'0070ddb7    8b9518ffffff            mov edx, dword ptr [ebp+ffffff18]
'0070ddbd    66899d14ffffff          mov word ptr [ebp+ffffff14], bx
'0070ddc4    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'0070ddca    894108                  mov dword ptr [ecx+08], eax
'0070ddcd    8b85f0feffff            mov eax, dword ptr [ebp+fffffef0]
'0070ddd3    89510c                  mov dword ptr [ecx+0c], edx
'0070ddd6    83ec10                  sub esp, 10
'0070ddd9    8bcc                    mov ecx, esp
'0070dddb    c785ecfeffff02000000    mov dword ptr [ebp+fffffeec], 00000002
'0070dde5    8b95ecfeffff            mov edx, dword ptr [ebp+fffffeec]
'0070ddeb    8911                    mov dword ptr [ecx], edx
'0070dded    8b95f8feffff            mov edx, dword ptr [ebp+fffffef8]
'0070ddf3    894104                  mov dword ptr [ecx+04], eax
'0070ddf6    b803000000              mov eax, 00000003
'0070ddfb    50                      push eax
'0070ddfc    894108                  mov dword ptr [ecx+08], eax
'0070ddff    8b06                    mov eax, dword ptr [esi]
'0070de01    689e000000              push 0000009e
'0070de06    56                      push esi
'0070de07    66899db4feffff          mov word ptr [ebp+fffffeb4], bx
'0070de0e    89510c                  mov dword ptr [ecx+0c], edx
'0070de11    ff9020030000            call dword ptr [eax+00000320]
'0070de17    50                      push eax
'0070de18    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070de1b    51                      push ecx
'0070de1c    ffd7                    call edi
Set var_43 = Nothing
'0070de1e    50                      push eax
'0070de1f    8d55ac                  lea edx, dword ptr [ebp-54]
'0070de22    52                      push edx

' *** Reference to "__vbaLateIdCallLd"
'0070de23    ff157c114000            call dword ptr [0040117c]
var_50 = var_43.UNK_-256 - 12_158
'0070de29    8b8dccfeffff            mov ecx, dword ptr [ebp+fffffecc]
'0070de2f    8b95d0feffff            mov edx, dword ptr [ebp+fffffed0]
'0070de35    83c440                  add esp, 40
'0070de38    50                      push eax
'0070de39    83ec10                  sub esp, 10
'0070de3c    8bc4                    mov eax, esp
'0070de3e    8908                    mov dword ptr [eax], ecx
'0070de40    8b8dd4feffff            mov ecx, dword ptr [ebp+fffffed4]
'0070de46    895004                  mov dword ptr [eax+04], edx
'0070de49    8b95d8feffff            mov edx, dword ptr [ebp+fffffed8]
'0070de4f    894808                  mov dword ptr [eax+08], ecx
'0070de52    89500c                  mov dword ptr [eax+0c], edx
'0070de55    8b95b0feffff            mov edx, dword ptr [ebp+fffffeb0]
'0070de5b    83ec10                  sub esp, 10
'0070de5e    8bcc                    mov ecx, esp
'0070de60    b802000000              mov eax, 00000002
'0070de65    8901                    mov dword ptr [ecx], eax
'0070de67    8b85b4feffff            mov eax, dword ptr [ebp+fffffeb4]
'0070de6d    895104                  mov dword ptr [ecx+04], edx
'0070de70    8b95b8feffff            mov edx, dword ptr [ebp+fffffeb8]
'0070de76    894108                  mov dword ptr [ecx+08], eax
'0070de79    89510c                  mov dword ptr [ecx+0c], edx
'0070de7c    8b9590feffff            mov edx, dword ptr [ebp+fffffe90]
'0070de82    83ec10                  sub esp, 10
'0070de85    8bcc                    mov ecx, esp
'0070de87    b802000000              mov eax, 00000002
'0070de8c    8901                    mov dword ptr [ecx], eax
'0070de8e    895104                  mov dword ptr [ecx+04], edx
'0070de91    b804000000              mov eax, 00000004
'0070de96    894108                  mov dword ptr [ecx+08], eax
'0070de99    8b8598feffff            mov eax, dword ptr [ebp+fffffe98]
'0070de9f    6a03                    push 03
'0070dea1    689e000000              push 0000009e
'0070dea6    89410c                  mov dword ptr [ecx+0c], eax
'0070dea9    8b0e                    mov ecx, dword ptr [esi]
'0070deab    56                      push esi
'0070deac    ff9120030000            call dword ptr [ecx+00000320]
'0070deb2    50                      push eax
'0070deb3    8d55c8                  lea edx, dword ptr [ebp-38]
'0070deb6    52                      push edx
'0070deb7    ffd7                    call edi
Set var_46 = Nothing
'0070deb9    50                      push eax
'0070deba    8d459c                  lea eax, dword ptr [ebp-64]
'0070debd    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'0070debe    ff157c114000            call dword ptr [0040117c]
var_51 = var_46.UNK_-256 - 12_158
'0070dec4    83c440                  add esp, 40
'0070dec7    50                      push eax

' *** Reference to "__vbaVarTstGe"
'0070dec8    ff15a8124000            call dword ptr [004012a8]
'0070dece    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0070ded1    51                      push ecx
'0070ded2    8d55cc                  lea edx, dword ptr [ebp-34]
'0070ded5    52                      push edx
'0070ded6    6a02                    push 02
'0070ded8    668985e0fdffff          mov word ptr [ebp+fffffde0], ax

' *** Reference to "__vbaFreeObjList"
'0070dedf    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46)
'0070dee5    8d459c                  lea eax, dword ptr [ebp-64]
'0070dee8    50                      push eax
'0070dee9    8d4dac                  lea ecx, dword ptr [ebp-54]
'0070deec    51                      push ecx
'0070deed    6a02                    push 02

' *** Reference to "__vbaFreeVarList"
'0070deef    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51)
'0070def5    83c418                  add esp, 18
'0070def8    6683bde0fdffff00        cmp word ptr [ebp+fffffde0], 00
'0070df00    0f84e4020000            je 70e1ea

If (((-256 - 12) >= (var_51))) Then
'0070df06    8b16                    mov edx, dword ptr [esi]
'0070df08    56                      push esi
'0070df09    ff920c030000            call dword ptr [edx+0000030c]
'0070df0f    50                      push eax
'0070df10    8d45c0                  lea eax, dword ptr [ebp-40]
'0070df13    50                      push eax
'0070df14    ffd7                    call edi
    Set var_5 = 
'0070df16    8b0e                    mov ecx, dword ptr [esi]
'0070df18    56                      push esi
'0070df19    8985d8fdffff            mov dword ptr [ebp+fffffdd8], eax
'0070df1f    ff910c030000            call dword ptr [ecx+0000030c]
'0070df25    50                      push eax
'0070df26    8d55cc                  lea edx, dword ptr [ebp-34]
'0070df29    52                      push edx
'0070df2a    ffd7                    call edi
    Set var_43 = var_5
'0070df2c    8b08                    mov ecx, dword ptr [eax]
'0070df2e    8d55e0                  lea edx, dword ptr [ebp-20]
'0070df31    52                      push edx
'0070df32    50                      push eax
'0070df33    8985e0fdffff            mov dword ptr [ebp+fffffde0], eax
'0070df39    ff91a0000000            call dword ptr [ecx+000000a0]
'0070df3f    dbe2                    fnclex
'0070df41    85c0                    test ax, ax
'0070df43    7d18                    jge 70df5d
    
    If (    var_43) Then
'0070df45    8b8de0fdffff            mov ecx, dword ptr [ebp+fffffde0]
'0070df4b    68a0000000              push 000000a0
'0070df50    68d00d4300              push 00430dd0
'0070df55    51                      push ecx
'0070df56    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070df57    ff1580104000            call dword ptr [00401080]
    
End If
'0070df5d    8b55e0                  mov edx, dword ptr [ebp-20]
'0070df60    52                      push edx
'0070df61    68ac654500              push 004565ac

' *** Reference to "__vbaStrCat"
'0070df66    ff1570104000            call dword ptr [00401070]
var_51 = (vbNullString) & ("Passage de ")
'0070df6c    8945b4                  mov dword ptr [ebp-4c], eax
'0070df6f    ba08000000              mov edx, 00000008
'0070df74    8955ac                  mov dword ptr [ebp-54], edx
'0070df77    8995ccfeffff            mov dword ptr [ebp+fffffecc], edx
'0070df7d    89954cfeffff            mov dword ptr [ebp+fffffe4c], edx
'0070df83    8b95d8fdffff            mov edx, dword ptr [ebp+fffffdd8]
'0070df89    66899d14ffffff          mov word ptr [ebp+ffffff14], bx
'0070df90    c785d4feffffc8654500    mov dword ptr [ebp+fffffed4], 004565c8
'0070df9a    c78564feffff01000000    mov dword ptr [ebp+fffffe64], 00000001
'0070dfa4    c7855cfeffff02000000    mov dword ptr [ebp+fffffe5c], 00000002
'0070dfae    c78554feffff44ed4300    mov dword ptr [ebp+fffffe54], 0043ed44
'0070dfb8    66899da4feffff          mov word ptr [ebp+fffffea4], bx
'0070dfbf    8b1a                    mov ebx, dword ptr [edx]
'0070dfc1    8d55ac                  lea edx, dword ptr [ebp-54]
'0070dfc4    52                      push edx
'0070dfc5    83ec10                  sub esp, 10
'0070dfc8    8bd4                    mov edx, esp
'0070dfca    33c0                    xor eax, eax
'0070dfcc    b903000000              mov ecx, 00000003
'0070dfd1    890a                    mov dword ptr [edx], ecx
'0070dfd3    8b8d30ffffff            mov ecx, dword ptr [ebp+ffffff30]
'0070dfd9    894a04                  mov dword ptr [edx+04], ecx
'0070dfdc    894208                  mov dword ptr [edx+08], eax
'0070dfdf    8985f4feffff            mov dword ptr [ebp+fffffef4], eax
'0070dfe5    8b8538ffffff            mov eax, dword ptr [ebp+ffffff38]
'0070dfeb    89420c                  mov dword ptr [edx+0c], eax
'0070dfee    8b9510ffffff            mov edx, dword ptr [ebp+ffffff10]
'0070dff4    83ec10                  sub esp, 10
'0070dff7    8bcc                    mov ecx, esp
'0070dff9    b802000000              mov eax, 00000002
'0070dffe    8901                    mov dword ptr [ecx], eax
'0070e000    8b8514ffffff            mov eax, dword ptr [ebp+ffffff14]
'0070e006    895104                  mov dword ptr [ecx+04], edx
'0070e009    8b9518ffffff            mov edx, dword ptr [ebp+ffffff18]
'0070e00f    894108                  mov dword ptr [ecx+08], eax
'0070e012    89510c                  mov dword ptr [ecx+0c], edx
'0070e015    8b95f0feffff            mov edx, dword ptr [ebp+fffffef0]
'0070e01b    83ec10                  sub esp, 10
'0070e01e    8bcc                    mov ecx, esp
'0070e020    b802000000              mov eax, 00000002
'0070e025    8901                    mov dword ptr [ecx], eax
'0070e027    8b85f4feffff            mov eax, dword ptr [ebp+fffffef4]
'0070e02d    895104                  mov dword ptr [ecx+04], edx
'0070e030    8b95f8feffff            mov edx, dword ptr [ebp+fffffef8]
'0070e036    6a03                    push 03
'0070e038    894108                  mov dword ptr [ecx+08], eax
'0070e03b    8b06                    mov eax, dword ptr [esi]
'0070e03d    689e000000              push 0000009e
'0070e042    56                      push esi
'0070e043    89510c                  mov dword ptr [ecx+0c], edx
'0070e046    ff9020030000            call dword ptr [eax+00000320]
'0070e04c    50                      push eax
'0070e04d    8d4dc8                  lea ecx, dword ptr [ebp-38]
'0070e050    51                      push ecx
'0070e051    ffd7                    call edi
Set var_46 = Nothing
'0070e053    50                      push eax
'0070e054    8d559c                  lea edx, dword ptr [ebp-64]
'0070e057    52                      push edx

' *** Reference to "__vbaLateIdCallLd"
'0070e058    ff157c114000            call dword ptr [0040117c]
var_51 = var_46.UNK_-256 - 12_158
'0070e05e    83c440                  add esp, 40
'0070e061    50                      push eax
'0070e062    8d458c                  lea eax, dword ptr [ebp-74]
'0070e065    50                      push eax

' *** Reference to "__vbaVarCat"
'0070e066    ff1508124000            call dword ptr [00401208]
'0070e06c    50                      push eax
'0070e06d    8d8dccfeffff            lea ecx, dword ptr [ebp+fffffecc]
'0070e073    51                      push ecx
'0070e074    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'0070e07a    52                      push edx

' *** Reference to "__vbaVarCat"
'0070e07b    ff1508124000            call dword ptr [00401208]
'0070e081    8b95c0feffff            mov edx, dword ptr [ebp+fffffec0]
'0070e087    50                      push eax
'0070e088    83ec10                  sub esp, 10
'0070e08b    8bcc                    mov ecx, esp
'0070e08d    b803000000              mov eax, 00000003
'0070e092    8901                    mov dword ptr [ecx], eax
'0070e094    895104                  mov dword ptr [ecx+04], edx
'0070e097    8b95a0feffff            mov edx, dword ptr [ebp+fffffea0]
'0070e09d    33c0                    xor eax, eax
var_num1 = Empty
'0070e09f    894108                  mov dword ptr [ecx+08], eax
'0070e0a2    8b85c8feffff            mov eax, dword ptr [ebp+fffffec8]
'0070e0a8    89410c                  mov dword ptr [ecx+0c], eax
'0070e0ab    83ec10                  sub esp, 10
'0070e0ae    8bcc                    mov ecx, esp
'0070e0b0    b802000000              mov eax, 00000002
'0070e0b5    8901                    mov dword ptr [ecx], eax
'0070e0b7    8b85a4feffff            mov eax, dword ptr [ebp+fffffea4]
'0070e0bd    895104                  mov dword ptr [ecx+04], edx
'0070e0c0    8b95a8feffff            mov edx, dword ptr [ebp+fffffea8]
'0070e0c6    894108                  mov dword ptr [ecx+08], eax
'0070e0c9    89510c                  mov dword ptr [ecx+0c], edx
'0070e0cc    8b9580feffff            mov edx, dword ptr [ebp+fffffe80]
'0070e0d2    83ec10                  sub esp, 10
'0070e0d5    8bcc                    mov ecx, esp
'0070e0d7    b802000000              mov eax, 00000002
'0070e0dc    8901                    mov dword ptr [ecx], eax
'0070e0de    895104                  mov dword ptr [ecx+04], edx
'0070e0e1    b801000000              mov eax, 00000001
'0070e0e6    894108                  mov dword ptr [ecx+08], eax
'0070e0e9    8b8588feffff            mov eax, dword ptr [ebp+fffffe88]
'0070e0ef    6a03                    push 03
'0070e0f1    689e000000              push 0000009e
'0070e0f6    89410c                  mov dword ptr [ecx+0c], eax
'0070e0f9    8b0e                    mov ecx, dword ptr [esi]
'0070e0fb    56                      push esi
'0070e0fc    ff9120030000            call dword ptr [ecx+00000320]
'0070e102    50                      push eax
'0070e103    8d55c4                  lea edx, dword ptr [ebp-3c]
'0070e106    52                      push edx
'0070e107    ffd7                    call edi
Set var_9 = Nothing
'0070e109    50                      push eax
'0070e10a    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'0070e110    50                      push eax

' *** Reference to "__vbaLateIdCallLd"
'0070e111    ff157c114000            call dword ptr [0040117c]
var_94 = var_9.UNK_-256 - 12_158
'0070e117    83c440                  add esp, 40
'0070e11a    50                      push eax
'0070e11b    8d8d5cfeffff            lea ecx, dword ptr [ebp+fffffe5c]
'0070e121    51                      push ecx
'0070e122    8d955cffffff            lea edx, dword ptr [ebp+ffffff5c]
'0070e128    52                      push edx

' *** Reference to "__vbaVarAdd"
'0070e129    ff158c124000            call dword ptr [0040128c]
'0070e12f    50                      push eax
'0070e130    8d854cffffff            lea eax, dword ptr [ebp+ffffff4c]
'0070e136    50                      push eax

' *** Reference to "__vbaVarCat"
'0070e137    ff1508124000            call dword ptr [00401208]
'0070e13d    50                      push eax
'0070e13e    8d8d4cfeffff            lea ecx, dword ptr [ebp+fffffe4c]
'0070e144    51                      push ecx
'0070e145    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'0070e14b    52                      push edx

' *** Reference to "__vbaVarCat"
'0070e14c    ff1508124000            call dword ptr [00401208]
'0070e152    50                      push eax
'0070e153    8d45dc                  lea eax, dword ptr [ebp-24]
'0070e156    50                      push eax

' *** Reference to "__vbaStrVarVal"
'0070e157    ff15fc114000            call dword ptr [004011fc]
'0070e15d    8bcb                    mov ecx, ebx
'0070e15f    8b9dd8fdffff            mov ebx, dword ptr [ebp+fffffdd8]
'0070e165    50                      push eax
'0070e166    53                      push ebx
'0070e167    ff91a4000000            call dword ptr [ecx+000000a4]
'0070e16d    dbe2                    fnclex
'0070e16f    85c0                    test ax, ax
'0070e171    7d12                    jge 70e185
'0070e173    68a4000000              push 000000a4
'0070e178    68d00d4300              push 00430dd0
'0070e17d    53                      push ebx
'0070e17e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070e17f    ff1580104000            call dword ptr [00401080]
'0070e185    8d55dc                  lea edx, dword ptr [ebp-24]
'0070e188    52                      push edx
'0070e189    8d45e0                  lea eax, dword ptr [ebp-20]
'0070e18c    50                      push eax
'0070e18d    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'0070e18f    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_46, 0)
'0070e195    8d4dc0                  lea ecx, dword ptr [ebp-40]
'0070e198    51                      push ecx
'0070e199    8d55c4                  lea edx, dword ptr [ebp-3c]
'0070e19c    52                      push edx
'0070e19d    8d45c8                  lea eax, dword ptr [ebp-38]
'0070e1a0    50                      push eax
'0070e1a1    8d4dcc                  lea ecx, dword ptr [ebp-34]
'0070e1a4    51                      push ecx
'0070e1a5    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'0070e1a7    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46, var_9, var_5)
'0070e1ad    8d953cffffff            lea edx, dword ptr [ebp+ffffff3c]
'0070e1b3    52                      push edx
'0070e1b4    8d854cffffff            lea eax, dword ptr [ebp+ffffff4c]
'0070e1ba    50                      push eax
'0070e1bb    8d8d5cffffff            lea ecx, dword ptr [ebp+ffffff5c]
'0070e1c1    51                      push ecx
'0070e1c2    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'0070e1c8    52                      push edx
'0070e1c9    8d856cffffff            lea eax, dword ptr [ebp+ffffff6c]
'0070e1cf    50                      push eax
'0070e1d0    8d4d8c                  lea ecx, dword ptr [ebp-74]
'0070e1d3    51                      push ecx
'0070e1d4    8d559c                  lea edx, dword ptr [ebp-64]
'0070e1d7    52                      push edx
'0070e1d8    8d45ac                  lea eax, dword ptr [ebp-54]
'0070e1db    50                      push eax
'0070e1dc    6a08                    push 08

' *** Reference to "__vbaFreeVarList"
'0070e1de    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_94, var_59, var_88, var_89, var_90)
'0070e1e4    8b5de8                  mov ebx, dword ptr [ebp-18]
'0070e1e7    83c444                  add esp, 44

'ERROR: Two many next close:
End If
'0070e1ea    8b45e4                  mov eax, dword ptr [ebp-1c]
'0070e1ed    8b08                    mov ecx, dword ptr [eax]
'0070e1ef    6a00                    push 00
'0070e1f1    6a01                    push 01
'0070e1f3    50                      push eax
'0070e1f4    ff9164010000            call dword ptr [ecx+00000164]
'0070e1fa    dbe2                    fnclex
'0070e1fc    85c0                    test ax, ax
'0070e1fe    7d15                    jge 70e215
'0070e200    8b55e4                  mov edx, dword ptr [ebp-1c]
'0070e203    6864010000              push 00000164
'0070e208    6830314300              push 00433130
'0070e20d    52                      push edx
'0070e20e    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070e20f    ff1580104000            call dword ptr [00401080]
'0070e215    b801000000              mov eax, 00000001
'0070e21a    6603c3                  add ax, bx
var_num1 = 1 + 1
'0070e21d    0f80e1000000            jo 70e304
'0070e223    8bd8                    mov ebx, eax
'0070e225    895de8                  mov dword ptr [ebp-18], ebx
'0070e228    e9adecffff              jmp 70ceda

'ERROR: Two many next close:
Loop
'0070e22d    8b45e4                  mov eax, dword ptr [ebp-1c]
'0070e230    8b08                    mov ecx, dword ptr [eax]
'0070e232    50                      push eax
'0070e233    ff91c4000000            call dword ptr [ecx+000000c4]
'0070e239    dbe2                    fnclex
'0070e23b    85c0                    test ax, ax
'0070e23d    7d15                    jge 70e254
'0070e23f    8b55e4                  mov edx, dword ptr [ebp-1c]
'0070e242    68c4000000              push 000000c4
'0070e247    6830314300              push 00433130
'0070e24c    52                      push edx
'0070e24d    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'0070e24e    ff1580104000            call dword ptr [00401080]
'0070e254    c745fc00000000          mov dword ptr [ebp-04], 00000000
'0070e25b    9b                      fwait
'0070e25c    68e0e27000              push 0070e2e0
'0070e261    eb73                    jmp 70e2d6
'0070e263    8d45d0                  lea eax, dword ptr [ebp-30]
'0070e266    50                      push eax
'0070e267    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'0070e26a    51                      push ecx
'0070e26b    8d55d8                  lea edx, dword ptr [ebp-28]
'0070e26e    52                      push edx
'0070e26f    8d45dc                  lea eax, dword ptr [ebp-24]
'0070e272    50                      push eax
'0070e273    8d4de0                  lea ecx, dword ptr [ebp-20]
'0070e276    51                      push ecx
'0070e277    6a05                    push 05

' *** Reference to "__vbaFreeStrList"
'0070e279    ff155c124000            call dword ptr [0040125c]
'#Cleanup( var_46, 0, 0, 0, -4680)
'0070e27f    8d55bc                  lea edx, dword ptr [ebp-44]
'0070e282    52                      push edx
'0070e283    8d45c0                  lea eax, dword ptr [ebp-40]
'0070e286    50                      push eax
'0070e287    8d4dc4                  lea ecx, dword ptr [ebp-3c]
'0070e28a    51                      push ecx
'0070e28b    8d55c8                  lea edx, dword ptr [ebp-38]
'0070e28e    52                      push edx
'0070e28f    8d45cc                  lea eax, dword ptr [ebp-34]
'0070e292    50                      push eax
'0070e293    6a05                    push 05

' *** Reference to "__vbaFreeObjList"
'0070e295    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_43, var_46, var_9, var_5, var_58)
'0070e29b    8d8d3cffffff            lea ecx, dword ptr [ebp+ffffff3c]
'0070e2a1    51                      push ecx
'0070e2a2    8d954cffffff            lea edx, dword ptr [ebp+ffffff4c]
'0070e2a8    52                      push edx
'0070e2a9    8d855cffffff            lea eax, dword ptr [ebp+ffffff5c]
'0070e2af    50                      push eax
'0070e2b0    8d8d6cffffff            lea ecx, dword ptr [ebp+ffffff6c]
'0070e2b6    51                      push ecx
'0070e2b7    8d957cffffff            lea edx, dword ptr [ebp+ffffff7c]
'0070e2bd    52                      push edx
'0070e2be    8d458c                  lea eax, dword ptr [ebp-74]
'0070e2c1    50                      push eax
'0070e2c2    8d4d9c                  lea ecx, dword ptr [ebp-64]
'0070e2c5    51                      push ecx
'0070e2c6    8d55ac                  lea edx, dword ptr [ebp-54]
'0070e2c9    52                      push edx
'0070e2ca    6a08                    push 08

' *** Reference to "__vbaFreeVarList"
'0070e2cc    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_50, var_51, var_52, var_59, var_94, var_88, var_89, var_90)
'0070e2d2    83c454                  add esp, 54
'0070e2d5    c3                      ret
'0070e2d6    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeObj"
'0070e2d9    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_40)
'0070e2df    c3                      ret
'0070e2e0    8b4508                  mov eax, dword ptr [ebp+08]
'0070e2e3    8b08                    mov ecx, dword ptr [eax]
'0070e2e5    50                      push eax
'0070e2e6    ff5108                  call dword ptr [ecx+08]
'0070e2e9    8b45fc                  mov eax, dword ptr [ebp-04]
'0070e2ec    8b4dec                  mov ecx, dword ptr [ebp-14]
'0070e2ef    5f                      pop edi
'0070e2f0    5e                      pop esi
    'Reference to '__except_list'
'0070e2f1    64890d00000000          mov dword ptr fs:[00000000], ecx
'0070e2f8    5b                      pop ebx
'0070e2f9    8be5                    mov esp, ebp
'0070e2fb    5d                      pop ebp
'0070e2fc    c20400                  ret 0004


End Sub


