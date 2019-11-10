VERSION 5.00

Begin VB.Form frmDescriptifCompetence
    Caption = "Description des compétences"
    ScaleMode = 1
    AutoRedraw = 0              'False
    FontTransparent = -1              'True
    BorderStyle = 1
    LinkTopic = "Form1"
    MaxButton = 0              'False
    MinButton = 0              'False
    ControlBox = 0              'False
    ClientLeft   = 45
    ClientTop    = 435
    ClientWidth  = 12945
    ClientHeight = 11865
    BeginProperty Font
        Name          = "Times New Roman"
        Size          = 9,75
        Charset       = 0
        Weight        = 700
        Underline     = 0              'False
        Italic        = 0              'False
        Strikethrough = 0              'False
    EndProperty
    Begin VB.TextBox fldstrinne
        BackColor = 12632256
        Left   = 1725
        Top    = 10425
        Width  = 11205
        Height = 1065
        TabIndex = 15
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
    Begin VB.TextBox fldstrRestriction
        BackColor = 12632256
        Left   = 1725
        Top    = 9300
        Width  = 11205
        Height = 1065
        TabIndex = 14
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
    Begin VB.TextBox fldstrSynergie
        BackColor = 12632256
        Left   = 1725
        Top    = 8220
        Width  = 11205
        Height = 1065
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
        Left   = 1725
        Top    = 7140
        Width  = 11205
        Height = 1065
        TabIndex = 12
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
    Begin VB.TextBox fldstrNouvelle
        BackColor = 12632256
        Left   = 1725
        Top    = 6060
        Width  = 11205
        Height = 1065
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
    Begin VB.TextBox fldstrAction
        BackColor = 12632256
        Left   = 1725
        Top    = 4980
        Width  = 11205
        Height = 1065
        TabIndex = 10
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
    Begin VB.CommandButton btnFermer
        Caption = "&Fermer"
        Left   = 11880
        Top    = 11520
        Width  = 1050
        Height = 330
        TabIndex = 3
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
    Begin VB.TextBox fldstrTest
        BackColor = 12632256
        Left   = 1725
        Top    = 405
        Width  = 11205
        Height = 4560
        TabIndex = 0
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
    Begin VB.Label Label7
        Caption = "Test inné"
        Left   = 80
        Top    = 10755
        Width  = 1680
        Height = 240
        TabIndex = 9
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
    Begin VB.Label Label6
        Caption = "Restrictions"
        Left   = 80
        Top    = 9675
        Width  = 1680
        Height = 240
        TabIndex = 8
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
    Begin VB.Label Label5
        Caption = "Synergie"
        Left   = 80
        Top    = 8595
        Width  = 1680
        Height = 240
        TabIndex = 7
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
    Begin VB.Label Label4
        Caption = "Spécial"
        Left   = 80
        Top    = 7515
        Width  = 1680
        Height = 240
        TabIndex = 6
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
    Begin VB.Label Label3
        Caption = "Nouvelles tentatives :"
        Left   = 80
        Top    = 6435
        Width  = 1680
        Height = 240
        TabIndex = 5
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
    Begin VB.Label Label1
        Caption = "Action"
        Left   = 80
        Top    = 5355
        Width  = 1680
        Height = 240
        TabIndex = 4
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
    Begin VB.Label labnom
        Left   = 45
        Top    = 45
        Width  = 12885
        Height = 330
        TabIndex = 2
        BorderStyle = 1
        Alignment = 2
    End
    Begin VB.Label Label2
        Caption = "Test de compétence"
        Left   = 75
        Top    = 2640
        Width  = 1680
        Height = 240
        TabIndex = 1
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
'Event for Form
Private Sub Form_Load()
'006cde90    55                      push ebp
'006cde91    8bec                    mov ebp, esp
'006cde93    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006cde96    6866724000              push 00407266
'006cde9b    64a100000000            mov ax, word ptr fs:[00000000]
'006cdea1    50                      push eax
    'Reference to '__except_list'
'006cdea2    64892500000000          mov dword ptr fs:[00000000], esp
'006cdea9    81ec80010000            sub esp, 00000180
'006cdeaf    53                      push ebx
'006cdeb0    56                      push esi
'006cdeb1    57                      push edi
'006cdeb2    8965f4                  mov dword ptr [ebp-0c], esp
'006cdeb5    c745f858684000          mov dword ptr [ebp-08], 00406858
'006cdebc    8b7d08                  mov edi, dword ptr [ebp+08]
'006cdebf    8bc7                    mov eax, edi
'006cdec1    83e001                  and eax, 01
'006cdec4    8945fc                  mov dword ptr [ebp-04], eax
'006cdec7    83e7fe                  and edi, fffffffe
'006cdeca    8b0f                    mov ecx, dword ptr [edi]
'006cdecc    57                      push edi
'006cdecd    897d08                  mov dword ptr [ebp+08], edi
'006cded0    ff5104                  call dword ptr [ecx+04]
'006cded3    8b17                    mov edx, dword ptr [edi]
'006cded5    33f6                    xor esi, esi
'006cded7    57                      push edi
'006cded8    8975e8                  mov dword ptr [ebp-18], esi
'006cdedb    8975e4                  mov dword ptr [ebp-1c], esi
'006cdede    8975e0                  mov dword ptr [ebp-20], esi
'006cdee1    8975dc                  mov dword ptr [ebp-24], esi
'006cdee4    8975d8                  mov dword ptr [ebp-28], esi
'006cdee7    8975d4                  mov dword ptr [ebp-2c], esi
'006cdeea    8975d0                  mov dword ptr [ebp-30], esi
'006cdeed    8975c0                  mov dword ptr [ebp-40], esi
'006cdef0    8975b0                  mov dword ptr [ebp-50], esi
'006cdef3    8975a0                  mov dword ptr [ebp-60], esi
'006cdef6    897580                  mov dword ptr [ebp-80], esi
'006cdef9    89b5ccfeffff            mov dword ptr [ebp+fffffecc], esi
'006cdeff    ff9234030000            call dword ptr [edx+00000334]
'006cdf05    50                      push eax
'006cdf06    8d45dc                  lea eax, dword ptr [ebp-24]
'006cdf09    50                      push eax

' *** Reference to "__vbaObjSet"
'006cdf0a    ff15b4104000            call dword ptr [004010b4]
Set var_39 = Me
'006cdf10    8b5734                  mov edx, dword ptr [edi+34]
'006cdf13    8bd8                    mov ebx, eax
'006cdf15    8b0b                    mov ecx, dword ptr [ebx]
'006cdf17    52                      push edx
'006cdf18    53                      push ebx
'006cdf19    ff5154                  call dword ptr [ecx+54]
'006cdf1c    dbe2                    fnclex
'006cdf1e    3bc6                    cmp eax, esi
'006cdf20    7d0f                    jge 6cdf31

If (var_39 < 0) Then
'006cdf22    6a54                    push 54
'006cdf24    683ce04300              push 0043e03c
'006cdf29    53                      push ebx
'006cdf2a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cdf2b    ff1580104000            call dword ptr [00401080]
    
End If
'006cdf31    8d4ddc                  lea ecx, dword ptr [ebp-24]

' *** Reference to "__vbaFreeObj"
'006cdf34    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'006cdf3a    8b3d4cb07200            mov edi, dword ptr [0072b04c]
'006cdf40    8d5ddc                  lea ebx, dword ptr [ebp-24]
'006cdf43    53                      push ebx
'006cdf44    83ec10                  sub esp, 10
'006cdf47    8bdc                    mov ebx, esp
'006cdf49    b90a000000              mov ecx, 0000000a
'006cdf4e    890b                    mov dword ptr [ebx], ecx
'006cdf50    8bf1                    mov esi, ecx
'006cdf52    8b8d74ffffff            mov ecx, dword ptr [ebp+ffffff74]
'006cdf58    894b04                  mov dword ptr [ebx+04], ecx
'006cdf5b    b804000280              mov eax, 80020004
'006cdf60    894308                  mov dword ptr [ebx+08], eax
'006cdf63    8bd0                    mov edx, eax
'006cdf65    8b857cffffff            mov eax, dword ptr [ebp+ffffff7c]
'006cdf6b    89430c                  mov dword ptr [ebx+0c], eax
'006cdf6e    8b4584                  mov eax, dword ptr [ebp-7c]
'006cdf71    83ec10                  sub esp, 10
'006cdf74    8bcc                    mov ecx, esp
'006cdf76    8931                    mov dword ptr [ecx], esi
'006cdf78    894104                  mov dword ptr [ecx+04], eax
'006cdf7b    895108                  mov dword ptr [ecx+08], edx
'006cdf7e    895588                  mov dword ptr [ebp-78], edx
'006cdf81    8b558c                  mov edx, dword ptr [ebp-74]
'006cdf84    89510c                  mov dword ptr [ecx+0c], edx
'006cdf87    8b5594                  mov edx, dword ptr [ebp-6c]
'006cdf8a    83ec10                  sub esp, 10
'006cdf8d    8bcc                    mov ecx, esp
'006cdf8f    b803000000              mov eax, 00000003
'006cdf94    8901                    mov dword ptr [ecx], eax
'006cdf96    895104                  mov dword ptr [ecx+04], edx
'006cdf99    8b559c                  mov edx, dword ptr [ebp-64]
'006cdf9c    c7459801000000          mov dword ptr [ebp-68], 00000001
'006cdfa3    8b4598                  mov eax, dword ptr [ebp-68]
'006cdfa6    894108                  mov dword ptr [ecx+08], eax
'006cdfa9    a14cb07200              mov ax, word ptr [0072b04c]
'006cdfae    6838284400              push 00442838
'006cdfb3    897580                  mov dword ptr [ebp-80], esi
'006cdfb6    8b3f                    mov edi, dword ptr [edi]
'006cdfb8    50                      push eax
'006cdfb9    89510c                  mov dword ptr [ecx+0c], edx
'006cdfbc    ff97bc000000            call dword ptr [edi+000000bc]
'006cdfc2    dbe2                    fnclex
'006cdfc4    85c0                    test ax, ax
'006cdfc6    7d18                    jge 6cdfe0
'006cdfc8    8b0d4cb07200            mov ecx, dword ptr [0072b04c]
'006cdfce    68bc000000              push 000000bc
'006cdfd3    68301f4300              push 00431f30
'006cdfd8    51                      push ecx
'006cdfd9    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cdfda    ff1580104000            call dword ptr [00401080]
'006cdfe0    8b45dc                  mov eax, dword ptr [ebp-24]
'006cdfe3    50                      push eax
'006cdfe4    8d55e8                  lea edx, dword ptr [ebp-18]
'006cdfe7    52                      push edx
'006cdfe8    c745dc00000000          mov dword ptr [ebp-24], 00000000

' *** Reference to "__vbaObjSet"
'006cdfef    ff15b4104000            call dword ptr [004010b4]
Set var_41 = var_39
'006cdff5    8b45e8                  mov eax, dword ptr [ebp-18]
'006cdff8    8b08                    mov ecx, dword ptr [eax]
'006cdffa    6830a84300              push 0043a830
'006cdfff    50                      push eax
'006ce000    ff5144                  call dword ptr [ecx+44]
'006ce003    dbe2                    fnclex
'006ce005    85c0                    test ax, ax
'006ce007    7d12                    jge 6ce01b
'006ce009    8b55e8                  mov edx, dword ptr [ebp-18]
'006ce00c    6a44                    push 44
'006ce00e    6830314300              push 00433130
'006ce013    52                      push edx
'006ce014    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006ce015    ff1580104000            call dword ptr [00401080]
'006ce01b    be0a000000              mov esi, 0000000a
'006ce020    83ec10                  sub esp, 10
'006ce023    8bdc                    mov ebx, esp
'006ce025    8bce                    mov ecx, esi
'006ce027    890b                    mov dword ptr [ebx], ecx
'006ce029    8b8dd4feffff            mov ecx, dword ptr [ebp+fffffed4]
'006ce02f    894b04                  mov dword ptr [ebx+04], ecx
'006ce032    ba04000280              mov edx, 80020004
'006ce037    8bc2                    mov eax, edx
'006ce039    894308                  mov dword ptr [ebx+08], eax
'006ce03c    8b85dcfeffff            mov eax, dword ptr [ebp+fffffedc]
'006ce042    89430c                  mov dword ptr [ebx+0c], eax
'006ce045    8b85e4feffff            mov eax, dword ptr [ebp+fffffee4]
'006ce04b    83ec10                  sub esp, 10
'006ce04e    8bcc                    mov ecx, esp
'006ce050    8931                    mov dword ptr [ecx], esi
'006ce052    894104                  mov dword ptr [ecx+04], eax
'006ce055    895108                  mov dword ptr [ecx+08], edx
'006ce058    895588                  mov dword ptr [ebp-78], edx
'006ce05b    8b95ecfeffff            mov edx, dword ptr [ebp+fffffeec]
'006ce061    89510c                  mov dword ptr [ecx+0c], edx
'006ce064    8b95f4feffff            mov edx, dword ptr [ebp+fffffef4]
'006ce06a    83ec10                  sub esp, 10
'006ce06d    8bcc                    mov ecx, esp
'006ce06f    8bc6                    mov eax, esi
'006ce071    8901                    mov dword ptr [ecx], eax
'006ce073    895104                  mov dword ptr [ecx+04], edx
'006ce076    8b9504ffffff            mov edx, dword ptr [ebp+ffffff04]
'006ce07c    b804000280              mov eax, 80020004
'006ce081    894108                  mov dword ptr [ecx+08], eax
'006ce084    8b85fcfeffff            mov eax, dword ptr [ebp+fffffefc]
'006ce08a    89410c                  mov dword ptr [ecx+0c], eax
'006ce08d    83ec10                  sub esp, 10
'006ce090    8bcc                    mov ecx, esp
'006ce092    8bc6                    mov eax, esi
'006ce094    8901                    mov dword ptr [ecx], eax
'006ce096    895104                  mov dword ptr [ecx+04], edx
'006ce099    8b9514ffffff            mov edx, dword ptr [ebp+ffffff14]
'006ce09f    b804000280              mov eax, 80020004
'006ce0a4    894108                  mov dword ptr [ecx+08], eax
'006ce0a7    8b850cffffff            mov eax, dword ptr [ebp+ffffff0c]
'006ce0ad    89410c                  mov dword ptr [ecx+0c], eax
'006ce0b0    83ec10                  sub esp, 10
'006ce0b3    8bcc                    mov ecx, esp
'006ce0b5    8bc6                    mov eax, esi
'006ce0b7    8901                    mov dword ptr [ecx], eax
'006ce0b9    895104                  mov dword ptr [ecx+04], edx
'006ce0bc    8b9524ffffff            mov edx, dword ptr [ebp+ffffff24]
'006ce0c2    8b7de8                  mov edi, dword ptr [ebp-18]
'006ce0c5    b804000280              mov eax, 80020004
'006ce0ca    894108                  mov dword ptr [ecx+08], eax
'006ce0cd    8b851cffffff            mov eax, dword ptr [ebp+ffffff1c]
'006ce0d3    89410c                  mov dword ptr [ecx+0c], eax
'006ce0d6    83ec10                  sub esp, 10
'006ce0d9    8bcc                    mov ecx, esp
'006ce0db    8bc6                    mov eax, esi
'006ce0dd    8901                    mov dword ptr [ecx], eax
'006ce0df    895104                  mov dword ptr [ecx+04], edx
'006ce0e2    8b9534ffffff            mov edx, dword ptr [ebp+ffffff34]
'006ce0e8    b804000280              mov eax, 80020004
'006ce0ed    894108                  mov dword ptr [ecx+08], eax
'006ce0f0    8b852cffffff            mov eax, dword ptr [ebp+ffffff2c]
'006ce0f6    89410c                  mov dword ptr [ecx+0c], eax
'006ce0f9    83ec10                  sub esp, 10
'006ce0fc    8bcc                    mov ecx, esp
'006ce0fe    8bc6                    mov eax, esi
'006ce100    8901                    mov dword ptr [ecx], eax
'006ce102    895104                  mov dword ptr [ecx+04], edx
'006ce105    8b9544ffffff            mov edx, dword ptr [ebp+ffffff44]
'006ce10b    b804000280              mov eax, 80020004
'006ce110    894108                  mov dword ptr [ecx+08], eax
'006ce113    8b853cffffff            mov eax, dword ptr [ebp+ffffff3c]
'006ce119    89410c                  mov dword ptr [ecx+0c], eax
'006ce11c    83ec10                  sub esp, 10
'006ce11f    8bc6                    mov eax, esi
'006ce121    8bcc                    mov ecx, esp
'006ce123    8901                    mov dword ptr [ecx], eax
'006ce125    897580                  mov dword ptr [ebp-80], esi
'006ce128    8b3f                    mov edi, dword ptr [edi]
'006ce12a    895104                  mov dword ptr [ecx+04], edx
'006ce12d    b804000280              mov eax, 80020004
'006ce132    894108                  mov dword ptr [ecx+08], eax
'006ce135    8b854cffffff            mov eax, dword ptr [ebp+ffffff4c]
'006ce13b    89410c                  mov dword ptr [ecx+0c], eax
'006ce13e    8b9554ffffff            mov edx, dword ptr [ebp+ffffff54]
'006ce144    83ec10                  sub esp, 10
'006ce147    8bcc                    mov ecx, esp
'006ce149    8bc6                    mov eax, esi
'006ce14b    8901                    mov dword ptr [ecx], eax
'006ce14d    895104                  mov dword ptr [ecx+04], edx
'006ce150    8b9564ffffff            mov edx, dword ptr [ebp+ffffff64]
'006ce156    83ec10                  sub esp, 10
'006ce159    b804000280              mov eax, 80020004
'006ce15e    894108                  mov dword ptr [ecx+08], eax
'006ce161    8b855cffffff            mov eax, dword ptr [ebp+ffffff5c]
'006ce167    89410c                  mov dword ptr [ecx+0c], eax
'006ce16a    8bcc                    mov ecx, esp
'006ce16c    8bc6                    mov eax, esi
'006ce16e    8901                    mov dword ptr [ecx], eax
'006ce170    895104                  mov dword ptr [ecx+04], edx
'006ce173    8b9574ffffff            mov edx, dword ptr [ebp+ffffff74]
'006ce179    8b5d08                  mov ebx, dword ptr [ebp+08]
'006ce17c    b804000280              mov eax, 80020004
'006ce181    894108                  mov dword ptr [ecx+08], eax
'006ce184    8b856cffffff            mov eax, dword ptr [ebp+ffffff6c]
'006ce18a    89410c                  mov dword ptr [ecx+0c], eax
'006ce18d    83ec10                  sub esp, 10
'006ce190    8bcc                    mov ecx, esp
'006ce192    8bc6                    mov eax, esi
'006ce194    8901                    mov dword ptr [ecx], eax
'006ce196    895104                  mov dword ptr [ecx+04], edx
'006ce199    b804000280              mov eax, 80020004
'006ce19e    894108                  mov dword ptr [ecx+08], eax
'006ce1a1    8b857cffffff            mov eax, dword ptr [ebp+ffffff7c]
'006ce1a7    89410c                  mov dword ptr [ecx+0c], eax
'006ce1aa    8b4584                  mov eax, dword ptr [ebp-7c]
'006ce1ad    83ec10                  sub esp, 10
'006ce1b0    8bcc                    mov ecx, esp
'006ce1b2    8bd6                    mov edx, esi
'006ce1b4    8911                    mov dword ptr [ecx], edx
'006ce1b6    8b5588                  mov edx, dword ptr [ebp-78]
'006ce1b9    894104                  mov dword ptr [ecx+04], eax
'006ce1bc    8b458c                  mov eax, dword ptr [ebp-74]
'006ce1bf    895108                  mov dword ptr [ecx+08], edx
'006ce1c2    8b5594                  mov edx, dword ptr [ebp-6c]
'006ce1c5    89410c                  mov dword ptr [ecx+0c], eax
'006ce1c8    83ec10                  sub esp, 10
'006ce1cb    8bcc                    mov ecx, esp
'006ce1cd    b808000000              mov eax, 00000008
'006ce1d2    8901                    mov dword ptr [ecx], eax
'006ce1d4    8b4334                  mov eax, dword ptr [ebx+34]
'006ce1d7    895104                  mov dword ptr [ecx+04], edx
'006ce1da    894108                  mov dword ptr [ecx+08], eax
'006ce1dd    8b459c                  mov eax, dword ptr [ebp-64]
'006ce1e0    89410c                  mov dword ptr [ecx+0c], eax
'006ce1e3    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ce1e6    68209e4300              push 00439e20
'006ce1eb    51                      push ecx
'006ce1ec    ff97f4000000            call dword ptr [edi+000000f4]
'006ce1f2    dbe2                    fnclex
'006ce1f4    85c0                    test ax, ax
'006ce1f6    7d19                    jge 6ce211
'006ce1f8    8b55e8                  mov edx, dword ptr [ebp-18]

' *** Reference to "__vbaHresultCheckObj"
'006ce1fb    8b3580104000            mov esi, dword ptr [00401080]
'006ce201    68f4000000              push 000000f4
'006ce206    6830314300              push 00433130
'006ce20b    52                      push edx
'006ce20c    50                      push eax
'006ce20d    ffd6                    call esi
'006ce20f    eb06                    jmp 6ce217

' *** Reference to "__vbaHresultCheckObj"
'006ce211    8b3580104000            mov esi, dword ptr [00401080]
'006ce217    8b45e8                  mov eax, dword ptr [ebp-18]
'006ce21a    8b08                    mov ecx, dword ptr [eax]
'006ce21c    8d95ccfeffff            lea edx, dword ptr [ebp+fffffecc]
'006ce222    52                      push edx
'006ce223    50                      push eax
'006ce224    ff515c                  call dword ptr [ecx+5c]
'006ce227    dbe2                    fnclex
'006ce229    85c0                    test ax, ax
'006ce22b    7d0e                    jge 6ce23b
'006ce22d    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ce230    6a5c                    push 5c
'006ce232    6830314300              push 00433130
'006ce237    51                      push ecx
'006ce238    50                      push eax
'006ce239    ffd6                    call esi
'006ce23b    6683bdccfeffff00        cmp word ptr [ebp+fffffecc], 00
'006ce243    0f853f110000            jne 6cf388

If (0 = 0) Then
'006ce249    8b13                    mov edx, dword ptr [ebx]
'006ce24b    53                      push ebx
'006ce24c    ff9234030000            call dword ptr [edx+00000334]

' *** Reference to "__vbaObjSet"
'006ce252    8b3db4104000            mov edi, dword ptr [004010b4]
'006ce258    50                      push eax
'006ce259    8d45d0                  lea eax, dword ptr [ebp-30]
'006ce25c    50                      push eax
'006ce25d    ffd7                    call edi
    Set var_4 = var_41
'006ce25f    8b0b                    mov ecx, dword ptr [ebx]
'006ce261    53                      push ebx
'006ce262    8985acfeffff            mov dword ptr [ebp+fffffeac], eax
'006ce268    ff9134030000            call dword ptr [ecx+00000334]
'006ce26e    50                      push eax
'006ce26f    8d55dc                  lea edx, dword ptr [ebp-24]
'006ce272    52                      push edx
'006ce273    ffd7                    call edi
    Set var_39 = var_4
'006ce275    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006ce278    8bf8                    mov edi, eax
'006ce27a    8b07                    mov eax, dword ptr [edi]
'006ce27c    51                      push ecx
'006ce27d    57                      push edi
'006ce27e    ff5050                  call dword ptr [eax+50]
'006ce281    dbe2                    fnclex
'006ce283    85c0                    test ax, ax
'006ce285    7d0b                    jge 6ce292
'006ce287    6a50                    push 50
'006ce289    683ce04300              push 0043e03c
'006ce28e    57                      push edi
'006ce28f    50                      push eax
'006ce290    ffd6                    call esi
'006ce292    8b55e4                  mov edx, dword ptr [ebp-1c]
'006ce295    52                      push edx
'006ce296    6814cf4300              push 0043cf14

' *** Reference to "__vbaStrCat"
'006ce29b    ff1570104000            call dword ptr [00401070]
    var_49 = (vbNullString) & (" (")
'006ce2a1    8945b8                  mov dword ptr [ebp-48], eax
'006ce2a4    8b45e8                  mov eax, dword ptr [ebp-18]
'006ce2a7    8d55d8                  lea edx, dword ptr [ebp-28]
'006ce2aa    52                      push edx
'006ce2ab    c745b008000000          mov dword ptr [ebp-50], 00000008
'006ce2b2    8b08                    mov ecx, dword ptr [eax]
'006ce2b4    50                      push eax
'006ce2b5    ff91b4000000            call dword ptr [ecx+000000b4]
'006ce2bb    dbe2                    fnclex
'006ce2bd    85c0                    test ax, ax
'006ce2bf    7d11                    jge 6ce2d2
'006ce2c1    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ce2c4    68b4000000              push 000000b4
'006ce2c9    6830314300              push 00433130
'006ce2ce    51                      push ecx
'006ce2cf    50                      push eax
'006ce2d0    ffd6                    call esi
'006ce2d2    8b45d8                  mov eax, dword ptr [ebp-28]
'006ce2d5    8b10                    mov edx, dword ptr [eax]
'006ce2d7    8d7dd4                  lea edi, dword ptr [ebp-2c]
'006ce2da    57                      push edi
'006ce2db    83ec10                  sub esp, 10
'006ce2de    8bfc                    mov edi, esp
'006ce2e0    b908000000              mov ecx, 00000008
'006ce2e5    890f                    mov dword ptr [edi], ecx
'006ce2e7    8b4d94                  mov ecx, dword ptr [ebp-6c]
'006ce2ea    894f04                  mov dword ptr [edi+04], ecx
'006ce2ed    b920254400              mov ecx, 00442520
'006ce2f2    894f08                  mov dword ptr [edi+08], ecx
'006ce2f5    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006ce2f8    50                      push eax
'006ce2f9    8985bcfeffff            mov dword ptr [ebp+fffffebc], eax
'006ce2ff    894f0c                  mov dword ptr [edi+0c], ecx
'006ce302    ff5230                  call dword ptr [edx+30]
'006ce305    dbe2                    fnclex
'006ce307    85c0                    test ax, ax
'006ce309    7d11                    jge 6ce31c
'006ce30b    8b95bcfeffff            mov edx, dword ptr [ebp+fffffebc]
'006ce311    6a30                    push 30
'006ce313    68d8304300              push 004330d8
'006ce318    52                      push edx
'006ce319    50                      push eax
'006ce31a    ffd6                    call esi
'006ce31c    8b45d4                  mov eax, dword ptr [ebp-2c]
'006ce31f    8b08                    mov ecx, dword ptr [eax]
'006ce321    8d55c0                  lea edx, dword ptr [ebp-40]
'006ce324    52                      push edx
'006ce325    50                      push eax
'006ce326    8bf8                    mov edi, eax
'006ce328    ff5144                  call dword ptr [ecx+44]
'006ce32b    dbe2                    fnclex
'006ce32d    85c0                    test ax, ax
'006ce32f    7d0b                    jge 6ce33c
'006ce331    6a44                    push 44
'006ce333    6880304300              push 00433080
'006ce338    57                      push edi
'006ce339    50                      push eax
'006ce33a    ffd6                    call esi
'006ce33c    8b85acfeffff            mov eax, dword ptr [ebp+fffffeac]
'006ce342    8b38                    mov edi, dword ptr [eax]
'006ce344    8d4db0                  lea ecx, dword ptr [ebp-50]
'006ce347    51                      push ecx
'006ce348    8d55c0                  lea edx, dword ptr [ebp-40]
'006ce34b    52                      push edx
'006ce34c    8d45a0                  lea eax, dword ptr [ebp-60]
'006ce34f    50                      push eax

' *** Reference to "__vbaVarCat"
'006ce350    ff1508124000            call dword ptr [00401208]
'006ce356    50                      push eax
'006ce357    8d4de0                  lea ecx, dword ptr [ebp-20]
'006ce35a    51                      push ecx

' *** Reference to "__vbaStrVarVal"
'006ce35b    ff15fc114000            call dword ptr [004011fc]
'006ce361    8bd7                    mov edx, edi
'006ce363    8bbdacfeffff            mov edi, dword ptr [ebp+fffffeac]
'006ce369    50                      push eax
'006ce36a    57                      push edi
'006ce36b    ff5254                  call dword ptr [edx+54]
'006ce36e    dbe2                    fnclex
'006ce370    85c0                    test ax, ax
'006ce372    7d0b                    jge 6ce37f
'006ce374    6a54                    push 54
'006ce376    683ce04300              push 0043e03c
'006ce37b    57                      push edi
'006ce37c    50                      push eax
'006ce37d    ffd6                    call esi
'006ce37f    8d45e0                  lea eax, dword ptr [ebp-20]
'006ce382    50                      push eax
'006ce383    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006ce386    51                      push ecx
'006ce387    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'006ce389    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( 0, 0)
'006ce38f    8d55d0                  lea edx, dword ptr [ebp-30]
'006ce392    52                      push edx
'006ce393    8d45d4                  lea eax, dword ptr [ebp-2c]
'006ce396    50                      push eax
'006ce397    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006ce39a    51                      push ecx
'006ce39b    8d55dc                  lea edx, dword ptr [ebp-24]
'006ce39e    52                      push edx
'006ce39f    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'006ce3a1    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_39, var_45, var_44, var_4)
'006ce3a7    8d45a0                  lea eax, dword ptr [ebp-60]
'006ce3aa    50                      push eax
'006ce3ab    8d4dc0                  lea ecx, dword ptr [ebp-40]
'006ce3ae    51                      push ecx
'006ce3af    8d55b0                  lea edx, dword ptr [ebp-50]
'006ce3b2    52                      push edx
'006ce3b3    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'006ce3b5    ff1540104000            call dword ptr [00401040]
    '#Cleanup( var_6, var_5, var_7)
'006ce3bb    8b45e8                  mov eax, dword ptr [ebp-18]
'006ce3be    8b08                    mov ecx, dword ptr [eax]
'006ce3c0    83c430                  add esp, 30
'006ce3c3    8d55dc                  lea edx, dword ptr [ebp-24]
'006ce3c6    52                      push edx
'006ce3c7    50                      push eax
'006ce3c8    ff91b4000000            call dword ptr [ecx+000000b4]
'006ce3ce    dbe2                    fnclex
'006ce3d0    85c0                    test ax, ax
'006ce3d2    7d11                    jge 6ce3e5
'006ce3d4    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ce3d7    68b4000000              push 000000b4
'006ce3dc    6830314300              push 00433130
'006ce3e1    51                      push ecx
'006ce3e2    50                      push eax
'006ce3e3    ffd6                    call esi
'006ce3e5    8b45dc                  mov eax, dword ptr [ebp-24]
'006ce3e8    8b10                    mov edx, dword ptr [eax]
'006ce3ea    8d7dd8                  lea edi, dword ptr [ebp-28]
'006ce3ed    57                      push edi
'006ce3ee    83ec10                  sub esp, 10
'006ce3f1    8bfc                    mov edi, esp
'006ce3f3    b908000000              mov ecx, 00000008
'006ce3f8    890f                    mov dword ptr [edi], ecx
'006ce3fa    8b4d94                  mov ecx, dword ptr [ebp-6c]
'006ce3fd    894f04                  mov dword ptr [edi+04], ecx
'006ce400    b920284400              mov ecx, 00442820
'006ce405    894f08                  mov dword ptr [edi+08], ecx
'006ce408    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006ce40b    50                      push eax
'006ce40c    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'006ce412    894f0c                  mov dword ptr [edi+0c], ecx
'006ce415    ff5230                  call dword ptr [edx+30]
'006ce418    dbe2                    fnclex
'006ce41a    85c0                    test ax, ax
'006ce41c    7d11                    jge 6ce42f
'006ce41e    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'006ce424    6a30                    push 30
'006ce426    68d8304300              push 004330d8
'006ce42b    52                      push edx
'006ce42c    50                      push eax
'006ce42d    ffd6                    call esi
'006ce42f    8b45d8                  mov eax, dword ptr [ebp-28]
'006ce432    8b08                    mov ecx, dword ptr [eax]
'006ce434    8d55c0                  lea edx, dword ptr [ebp-40]
'006ce437    52                      push edx
'006ce438    50                      push eax
'006ce439    8bf8                    mov edi, eax
'006ce43b    ff5144                  call dword ptr [ecx+44]
'006ce43e    dbe2                    fnclex
'006ce440    85c0                    test ax, ax
'006ce442    7d0b                    jge 6ce44f
'006ce444    6a44                    push 44
'006ce446    6880304300              push 00433080
'006ce44b    57                      push edi
'006ce44c    50                      push eax
'006ce44d    ffd6                    call esi
'006ce44f    8d45c0                  lea eax, dword ptr [ebp-40]
'006ce452    50                      push eax
'006ce453    8d4d80                  lea ecx, dword ptr [ebp-80]
'006ce456    51                      push ecx
'006ce457    c7458830284400          mov dword ptr [ebp-78], 00442830
'006ce45e    c7458008800000          mov dword ptr [ebp-80], 00008008

' *** Reference to "__vbaVarTstEq"
'006ce465    ff153c114000            call dword ptr [0040113c]
'006ce46b    8d55d8                  lea edx, dword ptr [ebp-28]
'006ce46e    668bf8                  mov di, ax
'006ce471    52                      push edx
'006ce472    8d45dc                  lea eax, dword ptr [ebp-24]
'006ce475    50                      push eax
'006ce476    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006ce478    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_39, var_45)
'006ce47e    83c40c                  add esp, 0c
'006ce481    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeVar"
'006ce484    ff1528104000            call dword ptr [00401028]
    '#Cleanup(var_5)
'006ce48a    6685ff                  test edi, edi
'006ce48d    0f84ac000000            je 6ce53f
    
    If (    ((var_5) = ("N"))) Then
'006ce493    8b0b                    mov ecx, dword ptr [ebx]
'006ce495    53                      push ebx
'006ce496    ff9134030000            call dword ptr [ecx+00000334]

' *** Reference to "__vbaObjSet"
'006ce49c    8b3db4104000            mov edi, dword ptr [004010b4]
'006ce4a2    50                      push eax
'006ce4a3    8d55d8                  lea edx, dword ptr [ebp-28]
'006ce4a6    52                      push edx
'006ce4a7    ffd7                    call edi
    Set var_45 = 
'006ce4a9    8985c0feffff            mov dword ptr [ebp+fffffec0], eax
'006ce4af    8b03                    mov eax, dword ptr [ebx]
'006ce4b1    53                      push ebx
'006ce4b2    ff9034030000            call dword ptr [eax+00000334]
'006ce4b8    50                      push eax
'006ce4b9    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006ce4bc    51                      push ecx
'006ce4bd    ffd7                    call edi
    Set var_39 = Nothing
'006ce4bf    8bf8                    mov edi, eax
'006ce4c1    8b17                    mov edx, dword ptr [edi]
'006ce4c3    8d45e4                  lea eax, dword ptr [ebp-1c]
'006ce4c6    50                      push eax
'006ce4c7    57                      push edi
'006ce4c8    ff5250                  call dword ptr [edx+50]
'006ce4cb    dbe2                    fnclex
'006ce4cd    85c0                    test ax, ax
'006ce4cf    7d0b                    jge 6ce4dc
'006ce4d1    6a50                    push 50
'006ce4d3    683ce04300              push 0043e03c
'006ce4d8    57                      push edi
'006ce4d9    50                      push eax
'006ce4da    ffd6                    call esi
'006ce4dc    8b55e4                  mov edx, dword ptr [ebp-1c]
'006ce4df    8b8dc0feffff            mov ecx, dword ptr [ebp+fffffec0]
'006ce4e5    8b39                    mov edi, dword ptr [ecx]
'006ce4e7    52                      push edx
'006ce4e8    68fc3a4500              push 00453afc

' *** Reference to "__vbaStrCat"
'006ce4ed    ff1570104000            call dword ptr [00401070]
    var_40 = (vbNullString) & ("; formation nécessaire")
'006ce4f3    8bd0                    mov edx, eax
'006ce4f5    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrMove"
'006ce4f8    ff15d0124000            call dword ptr [004012d0]
'006ce4fe    50                      push eax
'006ce4ff    8bc7                    mov eax, edi
'006ce501    8bbdc0feffff            mov edi, dword ptr [ebp+fffffec0]
'006ce507    57                      push edi
'006ce508    ff5054                  call dword ptr [eax+54]
'006ce50b    dbe2                    fnclex
'006ce50d    85c0                    test ax, ax
'006ce50f    7d0b                    jge 6ce51c
'006ce511    6a54                    push 54
'006ce513    683ce04300              push 0043e03c
'006ce518    57                      push edi
'006ce519    50                      push eax
'006ce51a    ffd6                    call esi
'006ce51c    8d4de0                  lea ecx, dword ptr [ebp-20]
'006ce51f    51                      push ecx
'006ce520    8d55e4                  lea edx, dword ptr [ebp-1c]
'006ce523    52                      push edx
'006ce524    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'006ce526    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( 0, -4524)
'006ce52c    8d45d8                  lea eax, dword ptr [ebp-28]
'006ce52f    50                      push eax
'006ce530    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006ce533    51                      push ecx
'006ce534    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006ce536    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_39, var_45)
'006ce53c    83c418                  add esp, 18
    
End If
'006ce53f    8b45e8                  mov eax, dword ptr [ebp-18]
'006ce542    8b10                    mov edx, dword ptr [eax]
'006ce544    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006ce547    51                      push ecx
'006ce548    50                      push eax
'006ce549    ff92b4000000            call dword ptr [edx+000000b4]
'006ce54f    dbe2                    fnclex
'006ce551    85c0                    test ax, ax
'006ce553    7d11                    jge 6ce566
'006ce555    8b55e8                  mov edx, dword ptr [ebp-18]
'006ce558    68b4000000              push 000000b4
'006ce55d    6830314300              push 00433130
'006ce562    52                      push edx
'006ce563    50                      push eax
'006ce564    ffd6                    call esi
'006ce566    8b45dc                  mov eax, dword ptr [ebp-24]
'006ce569    8b10                    mov edx, dword ptr [eax]
'006ce56b    8d7dd8                  lea edi, dword ptr [ebp-28]
'006ce56e    57                      push edi
'006ce56f    83ec10                  sub esp, 10
'006ce572    8bfc                    mov edi, esp
'006ce574    b908000000              mov ecx, 00000008
'006ce579    890f                    mov dword ptr [edi], ecx
'006ce57b    8b4d94                  mov ecx, dword ptr [ebp-6c]
'006ce57e    894f04                  mov dword ptr [edi+04], ecx
'006ce581    b934a64300              mov ecx, 0043a634
'006ce586    894f08                  mov dword ptr [edi+08], ecx
'006ce589    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006ce58c    50                      push eax
'006ce58d    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'006ce593    894f0c                  mov dword ptr [edi+0c], ecx
'006ce596    ff5230                  call dword ptr [edx+30]
'006ce599    dbe2                    fnclex
'006ce59b    85c0                    test ax, ax
'006ce59d    7d11                    jge 6ce5b0
'006ce59f    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'006ce5a5    6a30                    push 30
'006ce5a7    68d8304300              push 004330d8
'006ce5ac    52                      push edx
'006ce5ad    50                      push eax
'006ce5ae    ffd6                    call esi
'006ce5b0    8b45d8                  mov eax, dword ptr [ebp-28]
'006ce5b3    8945c8                  mov dword ptr [ebp-38], eax
'006ce5b6    8d45c0                  lea eax, dword ptr [ebp-40]
'006ce5b9    50                      push eax
'006ce5ba    c745d800000000          mov dword ptr [ebp-28], 00000000
'006ce5c1    c745c009000000          mov dword ptr [ebp-40], 00000009

' *** Reference to "__vbaBoolVarNull"
'006ce5c8    ff15fc104000            call dword ptr [004010fc]
'006ce5ce    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006ce5d1    668bf8                  mov di, ax

' *** Reference to "__vbaFreeObj"
'006ce5d4    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'006ce5da    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeVar"
'006ce5dd    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_5)
'006ce5e3    6685ff                  test edi, edi
'006ce5e6    0f84ac000000            je 6ce698

If (CBool(var_45)) Then
'006ce5ec    8b0b                    mov ecx, dword ptr [ebx]
'006ce5ee    53                      push ebx
'006ce5ef    ff9134030000            call dword ptr [ecx+00000334]

' *** Reference to "__vbaObjSet"
'006ce5f5    8b3db4104000            mov edi, dword ptr [004010b4]
'006ce5fb    50                      push eax
'006ce5fc    8d55d8                  lea edx, dword ptr [ebp-28]
'006ce5ff    52                      push edx
'006ce600    ffd7                    call edi
    Set var_45 = CBool(var_45)
'006ce602    8985c0feffff            mov dword ptr [ebp+fffffec0], eax
'006ce608    8b03                    mov eax, dword ptr [ebx]
'006ce60a    53                      push ebx
'006ce60b    ff9034030000            call dword ptr [eax+00000334]
'006ce611    50                      push eax
'006ce612    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006ce615    51                      push ecx
'006ce616    ffd7                    call edi
    Set var_39 = Nothing
'006ce618    8bf8                    mov edi, eax
'006ce61a    8b17                    mov edx, dword ptr [edi]
'006ce61c    8d45e4                  lea eax, dword ptr [ebp-1c]
'006ce61f    50                      push eax
'006ce620    57                      push edi
'006ce621    ff5250                  call dword ptr [edx+50]
'006ce624    dbe2                    fnclex
'006ce626    85c0                    test ax, ax
'006ce628    7d0b                    jge 6ce635
'006ce62a    6a50                    push 50
'006ce62c    683ce04300              push 0043e03c
'006ce631    57                      push edi
'006ce632    50                      push eax
'006ce633    ffd6                    call esi
'006ce635    8b55e4                  mov edx, dword ptr [ebp-1c]
'006ce638    8b8dc0feffff            mov ecx, dword ptr [ebp+fffffec0]
'006ce63e    8b39                    mov edi, dword ptr [ecx]
'006ce640    52                      push edx
'006ce641    68303b4500              push 00453b30

' *** Reference to "__vbaStrCat"
'006ce646    ff1570104000            call dword ptr [00401070]
    var_40 = (vbNullString) & ("; malus d'armure aux tests")
'006ce64c    8bd0                    mov edx, eax
'006ce64e    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrMove"
'006ce651    ff15d0124000            call dword ptr [004012d0]
'006ce657    50                      push eax
'006ce658    8bc7                    mov eax, edi
'006ce65a    8bbdc0feffff            mov edi, dword ptr [ebp+fffffec0]
'006ce660    57                      push edi
'006ce661    ff5054                  call dword ptr [eax+54]
'006ce664    dbe2                    fnclex
'006ce666    85c0                    test ax, ax
'006ce668    7d0b                    jge 6ce675
'006ce66a    6a54                    push 54
'006ce66c    683ce04300              push 0043e03c
'006ce671    57                      push edi
'006ce672    50                      push eax
'006ce673    ffd6                    call esi
'006ce675    8d4de0                  lea ecx, dword ptr [ebp-20]
'006ce678    51                      push ecx
'006ce679    8d55e4                  lea edx, dword ptr [ebp-1c]
'006ce67c    52                      push edx
'006ce67d    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'006ce67f    ff155c124000            call dword ptr [0040125c]
    '#Cleanup( 0, -4532)
'006ce685    8d45d8                  lea eax, dword ptr [ebp-28]
'006ce688    50                      push eax
'006ce689    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006ce68c    51                      push ecx
'006ce68d    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006ce68f    ff154c104000            call dword ptr [0040104c]
    '#Cleanup( var_39, var_45)
'006ce695    83c418                  add esp, 18
    
End If
'006ce698    8b13                    mov edx, dword ptr [ebx]
'006ce69a    53                      push ebx
'006ce69b    ff9234030000            call dword ptr [edx+00000334]

' *** Reference to "__vbaObjSet"
'006ce6a1    8b3db4104000            mov edi, dword ptr [004010b4]
'006ce6a7    50                      push eax
'006ce6a8    8d45d8                  lea eax, dword ptr [ebp-28]
'006ce6ab    50                      push eax
'006ce6ac    ffd7                    call edi
Set var_45 = CBool(var_45)
'006ce6ae    8b0b                    mov ecx, dword ptr [ebx]
'006ce6b0    53                      push ebx
'006ce6b1    8985c0feffff            mov dword ptr [ebp+fffffec0], eax
'006ce6b7    ff9134030000            call dword ptr [ecx+00000334]
'006ce6bd    50                      push eax
'006ce6be    8d55dc                  lea edx, dword ptr [ebp-24]
'006ce6c1    52                      push edx
'006ce6c2    ffd7                    call edi
Set var_39 = var_45
'006ce6c4    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006ce6c7    8bf8                    mov edi, eax
'006ce6c9    8b07                    mov eax, dword ptr [edi]
'006ce6cb    51                      push ecx
'006ce6cc    57                      push edi
'006ce6cd    ff5050                  call dword ptr [eax+50]
'006ce6d0    dbe2                    fnclex
'006ce6d2    85c0                    test ax, ax
'006ce6d4    7d0b                    jge 6ce6e1
'006ce6d6    6a50                    push 50
'006ce6d8    683ce04300              push 0043e03c
'006ce6dd    57                      push edi
'006ce6de    50                      push eax
'006ce6df    ffd6                    call esi
'006ce6e1    8b45e4                  mov eax, dword ptr [ebp-1c]
'006ce6e4    8b95c0feffff            mov edx, dword ptr [ebp+fffffec0]
'006ce6ea    8b3a                    mov edi, dword ptr [edx]
'006ce6ec    50                      push eax
'006ce6ed    682c384300              push 0043382c

' *** Reference to "__vbaStrCat"
'006ce6f2    ff1570104000            call dword ptr [00401070]
var_49 = (vbNullString) & (")")
'006ce6f8    8bd0                    mov edx, eax
'006ce6fa    8d4de0                  lea ecx, dword ptr [ebp-20]

' *** Reference to "__vbaStrMove"
'006ce6fd    ff15d0124000            call dword ptr [004012d0]
'006ce703    8bcf                    mov ecx, edi
'006ce705    8bbdc0feffff            mov edi, dword ptr [ebp+fffffec0]
'006ce70b    50                      push eax
'006ce70c    57                      push edi
'006ce70d    ff5154                  call dword ptr [ecx+54]
'006ce710    dbe2                    fnclex
'006ce712    85c0                    test ax, ax
'006ce714    7d0b                    jge 6ce721
'006ce716    6a54                    push 54
'006ce718    683ce04300              push 0043e03c
'006ce71d    57                      push edi
'006ce71e    50                      push eax
'006ce71f    ffd6                    call esi
'006ce721    8d55e0                  lea edx, dword ptr [ebp-20]
'006ce724    52                      push edx
'006ce725    8d45e4                  lea eax, dword ptr [ebp-1c]
'006ce728    50                      push eax
'006ce729    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'006ce72b    ff155c124000            call dword ptr [0040125c]
'#Cleanup( 0, -4536)
'006ce731    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006ce734    51                      push ecx
'006ce735    8d55dc                  lea edx, dword ptr [ebp-24]
'006ce738    52                      push edx
'006ce739    6a02                    push 02

' *** Reference to "__vbaFreeObjList"
'006ce73b    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45)
'006ce741    8b45e8                  mov eax, dword ptr [ebp-18]
'006ce744    8b08                    mov ecx, dword ptr [eax]
'006ce746    83c418                  add esp, 18
'006ce749    8d55dc                  lea edx, dword ptr [ebp-24]
'006ce74c    52                      push edx
'006ce74d    50                      push eax
'006ce74e    ff91b4000000            call dword ptr [ecx+000000b4]
'006ce754    dbe2                    fnclex
'006ce756    85c0                    test ax, ax
'006ce758    7d11                    jge 6ce76b
'006ce75a    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ce75d    68b4000000              push 000000b4
'006ce762    6830314300              push 00433130
'006ce767    51                      push ecx
'006ce768    50                      push eax
'006ce769    ffd6                    call esi
'006ce76b    8b45dc                  mov eax, dword ptr [ebp-24]
'006ce76e    8b10                    mov edx, dword ptr [eax]
'006ce770    8d7dd8                  lea edi, dword ptr [ebp-28]
'006ce773    57                      push edi
'006ce774    83ec10                  sub esp, 10
'006ce777    8bfc                    mov edi, esp
'006ce779    b908000000              mov ecx, 00000008
'006ce77e    890f                    mov dword ptr [edi], ecx
'006ce780    8b4d94                  mov ecx, dword ptr [ebp-6c]
'006ce783    894f04                  mov dword ptr [edi+04], ecx
'006ce786    b96c3b4500              mov ecx, 00453b6c
'006ce78b    894f08                  mov dword ptr [edi+08], ecx
'006ce78e    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006ce791    50                      push eax
'006ce792    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'006ce798    894f0c                  mov dword ptr [edi+0c], ecx
'006ce79b    ff5230                  call dword ptr [edx+30]
'006ce79e    dbe2                    fnclex
'006ce7a0    85c0                    test ax, ax
'006ce7a2    7d11                    jge 6ce7b5
'006ce7a4    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'006ce7aa    6a30                    push 30
'006ce7ac    68d8304300              push 004330d8
'006ce7b1    52                      push edx
'006ce7b2    50                      push eax
'006ce7b3    ffd6                    call esi
'006ce7b5    8b45d8                  mov eax, dword ptr [ebp-28]
'006ce7b8    8945c8                  mov dword ptr [ebp-38], eax
'006ce7bb    8d45c0                  lea eax, dword ptr [ebp-40]
'006ce7be    50                      push eax
'006ce7bf    c745d800000000          mov dword ptr [ebp-28], 00000000
'006ce7c6    c745c009000000          mov dword ptr [ebp-40], 00000009

' *** Reference to "rtcIsNull"
'006ce7cd    ff1540114000            call dword ptr [00401140]
'006ce7d3    33ff                    xor edi, edi
var_num7 = Empty
'006ce7d5    668bf8                  mov di, ax
'006ce7d8    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006ce7db    f7d7                    not edi

' *** Reference to "__vbaFreeObj"
'006ce7dd    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'006ce7e3    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeVar"
'006ce7e6    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_5)
'006ce7ec    6685ff                  test edi, edi
'006ce7ef    0f840d010000            je 6ce902

If (Not (IsNull(var_45))) Then
'006ce7f5    8b0b                    mov ecx, dword ptr [ebx]
'006ce7f7    53                      push ebx
'006ce7f8    ff9118030000            call dword ptr [ecx+00000318]
'006ce7fe    50                      push eax
'006ce7ff    8d55d4                  lea edx, dword ptr [ebp-2c]
'006ce802    52                      push edx

' *** Reference to "__vbaObjSet"
'006ce803    ff15b4104000            call dword ptr [004010b4]
    Set var_44 = IsNull(var_45)
'006ce809    8d55dc                  lea edx, dword ptr [ebp-24]
'006ce80c    8985b4feffff            mov dword ptr [ebp+fffffeb4], eax
'006ce812    8b45e8                  mov eax, dword ptr [ebp-18]
'006ce815    8b08                    mov ecx, dword ptr [eax]
'006ce817    52                      push edx
'006ce818    50                      push eax
'006ce819    ff91b4000000            call dword ptr [ecx+000000b4]
'006ce81f    dbe2                    fnclex
'006ce821    85c0                    test ax, ax
'006ce823    7d11                    jge 6ce836
'006ce825    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ce828    68b4000000              push 000000b4
'006ce82d    6830314300              push 00433130
'006ce832    51                      push ecx
'006ce833    50                      push eax
'006ce834    ffd6                    call esi
'006ce836    8b45dc                  mov eax, dword ptr [ebp-24]
'006ce839    8b38                    mov edi, dword ptr [eax]
'006ce83b    8d5dd8                  lea ebx, dword ptr [ebp-28]
'006ce83e    53                      push ebx
'006ce83f    83ec10                  sub esp, 10
'006ce842    8bdc                    mov ebx, esp
'006ce844    ba08000000              mov edx, 00000008
'006ce849    8913                    mov dword ptr [ebx], edx
'006ce84b    8b5594                  mov edx, dword ptr [ebp-6c]
'006ce84e    895304                  mov dword ptr [ebx+04], edx
'006ce851    b96c3b4500              mov ecx, 00453b6c
'006ce856    894b08                  mov dword ptr [ebx+08], ecx
'006ce859    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006ce85c    50                      push eax
'006ce85d    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'006ce863    894b0c                  mov dword ptr [ebx+0c], ecx
'006ce866    ff5730                  call dword ptr [edi+30]
'006ce869    dbe2                    fnclex
'006ce86b    85c0                    test ax, ax
'006ce86d    7d11                    jge 6ce880
'006ce86f    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'006ce875    6a30                    push 30
'006ce877    68d8304300              push 004330d8
'006ce87c    52                      push edx
'006ce87d    50                      push eax
'006ce87e    ffd6                    call esi
'006ce880    8b45d8                  mov eax, dword ptr [ebp-28]
'006ce883    8b08                    mov ecx, dword ptr [eax]
'006ce885    8d55c0                  lea edx, dword ptr [ebp-40]
'006ce888    52                      push edx
'006ce889    50                      push eax
'006ce88a    8bf8                    mov edi, eax
'006ce88c    ff5144                  call dword ptr [ecx+44]
'006ce88f    dbe2                    fnclex
'006ce891    85c0                    test ax, ax
'006ce893    7d0b                    jge 6ce8a0
    
    If (    var_45) Then
'006ce895    6a44                    push 44
'006ce897    6880304300              push 00433080
'006ce89c    57                      push edi
'006ce89d    50                      push eax
'006ce89e    ffd6                    call esi
    
End If
'006ce8a0    8bbdb4feffff            mov edi, dword ptr [ebp+fffffeb4]
'006ce8a6    8b1f                    mov ebx, dword ptr [edi]
'006ce8a8    8d45c0                  lea eax, dword ptr [ebp-40]
'006ce8ab    50                      push eax

' *** Reference to "__vbaStrVarMove"
'006ce8ac    ff153c104000            call dword ptr [0040103c]
'006ce8b2    8bd0                    mov edx, eax
'006ce8b4    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006ce8b7    ff15d0124000            call dword ptr [004012d0]
'006ce8bd    50                      push eax
'006ce8be    57                      push edi
'006ce8bf    ff93a4000000            call dword ptr [ebx+000000a4]
'006ce8c5    dbe2                    fnclex
'006ce8c7    85c0                    test ax, ax
'006ce8c9    7d0e                    jge 6ce8d9
'006ce8cb    68a4000000              push 000000a4
'006ce8d0    68d00d4300              push 00430dd0
'006ce8d5    57                      push edi
'006ce8d6    50                      push eax
'006ce8d7    ffd6                    call esi
'006ce8d9    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006ce8dc    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006ce8e2    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006ce8e5    51                      push ecx
'006ce8e6    8d55d8                  lea edx, dword ptr [ebp-28]
'006ce8e9    52                      push edx
'006ce8ea    8d45dc                  lea eax, dword ptr [ebp-24]
'006ce8ed    50                      push eax
'006ce8ee    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006ce8f0    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45, var_44)
'006ce8f6    83c410                  add esp, 10
'006ce8f9    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeVar"
'006ce8fc    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_5)
'ERROR: Two many next close:
End If
'006ce902    8b45e8                  mov eax, dword ptr [ebp-18]
'006ce905    8b08                    mov ecx, dword ptr [eax]
'006ce907    8d55dc                  lea edx, dword ptr [ebp-24]
'006ce90a    52                      push edx
'006ce90b    50                      push eax
'006ce90c    ff91b4000000            call dword ptr [ecx+000000b4]
'006ce912    dbe2                    fnclex
'006ce914    85c0                    test ax, ax
'006ce916    7d11                    jge 6ce929
'006ce918    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ce91b    68b4000000              push 000000b4
'006ce920    6830314300              push 00433130
'006ce925    51                      push ecx
'006ce926    50                      push eax
'006ce927    ffd6                    call esi
'006ce929    8b45dc                  mov eax, dword ptr [ebp-24]
'006ce92c    8b38                    mov edi, dword ptr [eax]
'006ce92e    8d5dd8                  lea ebx, dword ptr [ebp-28]
'006ce931    53                      push ebx
'006ce932    83ec10                  sub esp, 10
'006ce935    8bdc                    mov ebx, esp
'006ce937    ba08000000              mov edx, 00000008
'006ce93c    8913                    mov dword ptr [ebx], edx
'006ce93e    8b5594                  mov edx, dword ptr [ebp-6c]
'006ce941    895304                  mov dword ptr [ebx+04], edx
'006ce944    b97c3b4500              mov ecx, 00453b7c
'006ce949    894b08                  mov dword ptr [ebx+08], ecx
'006ce94c    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006ce94f    50                      push eax
'006ce950    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'006ce956    894b0c                  mov dword ptr [ebx+0c], ecx
'006ce959    ff5730                  call dword ptr [edi+30]
'006ce95c    dbe2                    fnclex
'006ce95e    85c0                    test ax, ax
'006ce960    7d11                    jge 6ce973
'006ce962    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'006ce968    6a30                    push 30
'006ce96a    68d8304300              push 004330d8
'006ce96f    52                      push edx
'006ce970    50                      push eax
'006ce971    ffd6                    call esi
'006ce973    8b45d8                  mov eax, dword ptr [ebp-28]
'006ce976    8945c8                  mov dword ptr [ebp-38], eax
'006ce979    8d45c0                  lea eax, dword ptr [ebp-40]
'006ce97c    50                      push eax
'006ce97d    c745d800000000          mov dword ptr [ebp-28], 00000000
'006ce984    c745c009000000          mov dword ptr [ebp-40], 00000009

' *** Reference to "rtcIsNull"
'006ce98b    ff1540114000            call dword ptr [00401140]
'006ce991    33ff                    xor edi, edi
var_num7 = Empty
'006ce993    668bf8                  mov di, ax
'006ce996    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006ce999    f7d7                    not edi

' *** Reference to "__vbaFreeObj"
'006ce99b    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'006ce9a1    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeVar"
'006ce9a4    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_5)
'006ce9aa    6685ff                  test edi, edi
'006ce9ad    0f8410010000            je 6ceac3

If (Not (IsNull(var_45))) Then
'006ce9b3    8b4508                  mov eax, dword ptr [ebp+08]
'006ce9b6    8b08                    mov ecx, dword ptr [eax]
'006ce9b8    50                      push eax
'006ce9b9    ff9110030000            call dword ptr [ecx+00000310]
'006ce9bf    50                      push eax
'006ce9c0    8d55d4                  lea edx, dword ptr [ebp-2c]
'006ce9c3    52                      push edx

' *** Reference to "__vbaObjSet"
'006ce9c4    ff15b4104000            call dword ptr [004010b4]
    Set var_44 = Me
'006ce9ca    8d55dc                  lea edx, dword ptr [ebp-24]
'006ce9cd    8985b4feffff            mov dword ptr [ebp+fffffeb4], eax
'006ce9d3    8b45e8                  mov eax, dword ptr [ebp-18]
'006ce9d6    8b08                    mov ecx, dword ptr [eax]
'006ce9d8    52                      push edx
'006ce9d9    50                      push eax
'006ce9da    ff91b4000000            call dword ptr [ecx+000000b4]
'006ce9e0    dbe2                    fnclex
'006ce9e2    85c0                    test ax, ax
'006ce9e4    7d11                    jge 6ce9f7
'006ce9e6    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ce9e9    68b4000000              push 000000b4
'006ce9ee    6830314300              push 00433130
'006ce9f3    51                      push ecx
'006ce9f4    50                      push eax
'006ce9f5    ffd6                    call esi
'006ce9f7    8b45dc                  mov eax, dword ptr [ebp-24]
'006ce9fa    8b38                    mov edi, dword ptr [eax]
'006ce9fc    8d5dd8                  lea ebx, dword ptr [ebp-28]
'006ce9ff    53                      push ebx
'006cea00    83ec10                  sub esp, 10
'006cea03    8bdc                    mov ebx, esp
'006cea05    ba08000000              mov edx, 00000008
'006cea0a    8913                    mov dword ptr [ebx], edx
'006cea0c    8b5594                  mov edx, dword ptr [ebp-6c]
'006cea0f    895304                  mov dword ptr [ebx+04], edx
'006cea12    b97c3b4500              mov ecx, 00453b7c
'006cea17    894b08                  mov dword ptr [ebx+08], ecx
'006cea1a    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006cea1d    50                      push eax
'006cea1e    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'006cea24    894b0c                  mov dword ptr [ebx+0c], ecx
'006cea27    ff5730                  call dword ptr [edi+30]
'006cea2a    dbe2                    fnclex
'006cea2c    85c0                    test ax, ax
'006cea2e    7d11                    jge 6cea41
'006cea30    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'006cea36    6a30                    push 30
'006cea38    68d8304300              push 004330d8
'006cea3d    52                      push edx
'006cea3e    50                      push eax
'006cea3f    ffd6                    call esi
'006cea41    8b45d8                  mov eax, dword ptr [ebp-28]
'006cea44    8b08                    mov ecx, dword ptr [eax]
'006cea46    8d55c0                  lea edx, dword ptr [ebp-40]
'006cea49    52                      push edx
'006cea4a    50                      push eax
'006cea4b    8bf8                    mov edi, eax
'006cea4d    ff5144                  call dword ptr [ecx+44]
'006cea50    dbe2                    fnclex
'006cea52    85c0                    test ax, ax
'006cea54    7d0b                    jge 6cea61
    
    If (    var_45) Then
'006cea56    6a44                    push 44
'006cea58    6880304300              push 00433080
'006cea5d    57                      push edi
'006cea5e    50                      push eax
'006cea5f    ffd6                    call esi
    
End If
'006cea61    8bbdb4feffff            mov edi, dword ptr [ebp+fffffeb4]
'006cea67    8b1f                    mov ebx, dword ptr [edi]
'006cea69    8d45c0                  lea eax, dword ptr [ebp-40]
'006cea6c    50                      push eax

' *** Reference to "__vbaStrVarMove"
'006cea6d    ff153c104000            call dword ptr [0040103c]
'006cea73    8bd0                    mov edx, eax
'006cea75    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006cea78    ff15d0124000            call dword ptr [004012d0]
'006cea7e    50                      push eax
'006cea7f    57                      push edi
'006cea80    ff93a4000000            call dword ptr [ebx+000000a4]
'006cea86    dbe2                    fnclex
'006cea88    85c0                    test ax, ax
'006cea8a    7d0e                    jge 6cea9a
'006cea8c    68a4000000              push 000000a4
'006cea91    68d00d4300              push 00430dd0
'006cea96    57                      push edi
'006cea97    50                      push eax
'006cea98    ffd6                    call esi
'006cea9a    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006cea9d    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006ceaa3    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006ceaa6    51                      push ecx
'006ceaa7    8d55d8                  lea edx, dword ptr [ebp-28]
'006ceaaa    52                      push edx
'006ceaab    8d45dc                  lea eax, dword ptr [ebp-24]
'006ceaae    50                      push eax
'006ceaaf    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006ceab1    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45, var_44)
'006ceab7    83c410                  add esp, 10
'006ceaba    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeVar"
'006ceabd    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_5)
'ERROR: Two many next close:
End If
'006ceac3    8b45e8                  mov eax, dword ptr [ebp-18]
'006ceac6    8b08                    mov ecx, dword ptr [eax]
'006ceac8    8d55dc                  lea edx, dword ptr [ebp-24]
'006ceacb    52                      push edx
'006ceacc    50                      push eax
'006ceacd    ff91b4000000            call dword ptr [ecx+000000b4]
'006cead3    dbe2                    fnclex
'006cead5    85c0                    test ax, ax
'006cead7    7d11                    jge 6ceaea
'006cead9    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ceadc    68b4000000              push 000000b4
'006ceae1    6830314300              push 00433130
'006ceae6    51                      push ecx
'006ceae7    50                      push eax
'006ceae8    ffd6                    call esi
'006ceaea    8b45dc                  mov eax, dword ptr [ebp-24]
'006ceaed    8b38                    mov edi, dword ptr [eax]
'006ceaef    8d5dd8                  lea ebx, dword ptr [ebp-28]
'006ceaf2    53                      push ebx
'006ceaf3    83ec10                  sub esp, 10
'006ceaf6    8bdc                    mov ebx, esp
'006ceaf8    ba08000000              mov edx, 00000008
'006ceafd    8913                    mov dword ptr [ebx], edx
'006ceaff    8b5594                  mov edx, dword ptr [ebp-6c]
'006ceb02    895304                  mov dword ptr [ebx+04], edx
'006ceb05    b9903b4500              mov ecx, 00453b90
'006ceb0a    894b08                  mov dword ptr [ebx+08], ecx
'006ceb0d    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006ceb10    50                      push eax
'006ceb11    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'006ceb17    894b0c                  mov dword ptr [ebx+0c], ecx
'006ceb1a    ff5730                  call dword ptr [edi+30]
'006ceb1d    dbe2                    fnclex
'006ceb1f    85c0                    test ax, ax
'006ceb21    7d11                    jge 6ceb34
'006ceb23    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'006ceb29    6a30                    push 30
'006ceb2b    68d8304300              push 004330d8
'006ceb30    52                      push edx
'006ceb31    50                      push eax
'006ceb32    ffd6                    call esi
'006ceb34    8b45d8                  mov eax, dword ptr [ebp-28]
'006ceb37    8945c8                  mov dword ptr [ebp-38], eax
'006ceb3a    8d45c0                  lea eax, dword ptr [ebp-40]
'006ceb3d    50                      push eax
'006ceb3e    c745d800000000          mov dword ptr [ebp-28], 00000000
'006ceb45    c745c009000000          mov dword ptr [ebp-40], 00000009

' *** Reference to "rtcIsNull"
'006ceb4c    ff1540114000            call dword ptr [00401140]
'006ceb52    33ff                    xor edi, edi
var_num7 = Empty
'006ceb54    668bf8                  mov di, ax
'006ceb57    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006ceb5a    f7d7                    not edi

' *** Reference to "__vbaFreeObj"
'006ceb5c    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'006ceb62    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeVar"
'006ceb65    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_5)
'006ceb6b    6685ff                  test edi, edi
'006ceb6e    0f8410010000            je 6cec84

If (Not (IsNull(var_45))) Then
'006ceb74    8b4508                  mov eax, dword ptr [ebp+08]
'006ceb77    8b08                    mov ecx, dword ptr [eax]
'006ceb79    50                      push eax
'006ceb7a    ff910c030000            call dword ptr [ecx+0000030c]
'006ceb80    50                      push eax
'006ceb81    8d55d4                  lea edx, dword ptr [ebp-2c]
'006ceb84    52                      push edx

' *** Reference to "__vbaObjSet"
'006ceb85    ff15b4104000            call dword ptr [004010b4]
    Set var_44 = Me
'006ceb8b    8d55dc                  lea edx, dword ptr [ebp-24]
'006ceb8e    8985b4feffff            mov dword ptr [ebp+fffffeb4], eax
'006ceb94    8b45e8                  mov eax, dword ptr [ebp-18]
'006ceb97    8b08                    mov ecx, dword ptr [eax]
'006ceb99    52                      push edx
'006ceb9a    50                      push eax
'006ceb9b    ff91b4000000            call dword ptr [ecx+000000b4]
'006ceba1    dbe2                    fnclex
'006ceba3    85c0                    test ax, ax
'006ceba5    7d11                    jge 6cebb8
'006ceba7    8b4de8                  mov ecx, dword ptr [ebp-18]
'006cebaa    68b4000000              push 000000b4
'006cebaf    6830314300              push 00433130
'006cebb4    51                      push ecx
'006cebb5    50                      push eax
'006cebb6    ffd6                    call esi
'006cebb8    8b45dc                  mov eax, dword ptr [ebp-24]
'006cebbb    8b38                    mov edi, dword ptr [eax]
'006cebbd    8d5dd8                  lea ebx, dword ptr [ebp-28]
'006cebc0    53                      push ebx
'006cebc1    83ec10                  sub esp, 10
'006cebc4    8bdc                    mov ebx, esp
'006cebc6    ba08000000              mov edx, 00000008
'006cebcb    8913                    mov dword ptr [ebx], edx
'006cebcd    8b5594                  mov edx, dword ptr [ebp-6c]
'006cebd0    895304                  mov dword ptr [ebx+04], edx
'006cebd3    b9903b4500              mov ecx, 00453b90
'006cebd8    894b08                  mov dword ptr [ebx+08], ecx
'006cebdb    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006cebde    50                      push eax
'006cebdf    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'006cebe5    894b0c                  mov dword ptr [ebx+0c], ecx
'006cebe8    ff5730                  call dword ptr [edi+30]
'006cebeb    dbe2                    fnclex
'006cebed    85c0                    test ax, ax
'006cebef    7d11                    jge 6cec02
'006cebf1    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'006cebf7    6a30                    push 30
'006cebf9    68d8304300              push 004330d8
'006cebfe    52                      push edx
'006cebff    50                      push eax
'006cec00    ffd6                    call esi
'006cec02    8b45d8                  mov eax, dword ptr [ebp-28]
'006cec05    8b08                    mov ecx, dword ptr [eax]
'006cec07    8d55c0                  lea edx, dword ptr [ebp-40]
'006cec0a    52                      push edx
'006cec0b    50                      push eax
'006cec0c    8bf8                    mov edi, eax
'006cec0e    ff5144                  call dword ptr [ecx+44]
'006cec11    dbe2                    fnclex
'006cec13    85c0                    test ax, ax
'006cec15    7d0b                    jge 6cec22
    
    If (    var_45) Then
'006cec17    6a44                    push 44
'006cec19    6880304300              push 00433080
'006cec1e    57                      push edi
'006cec1f    50                      push eax
'006cec20    ffd6                    call esi
    
End If
'006cec22    8bbdb4feffff            mov edi, dword ptr [ebp+fffffeb4]
'006cec28    8b1f                    mov ebx, dword ptr [edi]
'006cec2a    8d45c0                  lea eax, dword ptr [ebp-40]
'006cec2d    50                      push eax

' *** Reference to "__vbaStrVarMove"
'006cec2e    ff153c104000            call dword ptr [0040103c]
'006cec34    8bd0                    mov edx, eax
'006cec36    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006cec39    ff15d0124000            call dword ptr [004012d0]
'006cec3f    50                      push eax
'006cec40    57                      push edi
'006cec41    ff93a4000000            call dword ptr [ebx+000000a4]
'006cec47    dbe2                    fnclex
'006cec49    85c0                    test ax, ax
'006cec4b    7d0e                    jge 6cec5b
'006cec4d    68a4000000              push 000000a4
'006cec52    68d00d4300              push 00430dd0
'006cec57    57                      push edi
'006cec58    50                      push eax
'006cec59    ffd6                    call esi
'006cec5b    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006cec5e    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006cec64    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006cec67    51                      push ecx
'006cec68    8d55d8                  lea edx, dword ptr [ebp-28]
'006cec6b    52                      push edx
'006cec6c    8d45dc                  lea eax, dword ptr [ebp-24]
'006cec6f    50                      push eax
'006cec70    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006cec72    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45, var_44)
'006cec78    83c410                  add esp, 10
'006cec7b    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeVar"
'006cec7e    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_5)
'ERROR: Two many next close:
End If
'006cec84    8b45e8                  mov eax, dword ptr [ebp-18]
'006cec87    8b08                    mov ecx, dword ptr [eax]
'006cec89    8d55dc                  lea edx, dword ptr [ebp-24]
'006cec8c    52                      push edx
'006cec8d    50                      push eax
'006cec8e    ff91b4000000            call dword ptr [ecx+000000b4]
'006cec94    dbe2                    fnclex
'006cec96    85c0                    test ax, ax
'006cec98    7d11                    jge 6cecab
'006cec9a    8b4de8                  mov ecx, dword ptr [ebp-18]
'006cec9d    68b4000000              push 000000b4
'006ceca2    6830314300              push 00433130
'006ceca7    51                      push ecx
'006ceca8    50                      push eax
'006ceca9    ffd6                    call esi
'006cecab    8b45dc                  mov eax, dword ptr [ebp-24]
'006cecae    8b38                    mov edi, dword ptr [eax]
'006cecb0    8d5dd8                  lea ebx, dword ptr [ebp-28]
'006cecb3    53                      push ebx
'006cecb4    83ec10                  sub esp, 10
'006cecb7    8bdc                    mov ebx, esp
'006cecb9    ba08000000              mov edx, 00000008
'006cecbe    8913                    mov dword ptr [ebx], edx
'006cecc0    8b5594                  mov edx, dword ptr [ebp-6c]
'006cecc3    895304                  mov dword ptr [ebx+04], edx
'006cecc6    b944384500              mov ecx, 00453844
'006ceccb    894b08                  mov dword ptr [ebx+08], ecx
'006cecce    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006cecd1    50                      push eax
'006cecd2    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'006cecd8    894b0c                  mov dword ptr [ebx+0c], ecx
'006cecdb    ff5730                  call dword ptr [edi+30]
'006cecde    dbe2                    fnclex
'006cece0    85c0                    test ax, ax
'006cece2    7d11                    jge 6cecf5
'006cece4    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'006cecea    6a30                    push 30
'006cecec    68d8304300              push 004330d8
'006cecf1    52                      push edx
'006cecf2    50                      push eax
'006cecf3    ffd6                    call esi
'006cecf5    8b45d8                  mov eax, dword ptr [ebp-28]
'006cecf8    8945c8                  mov dword ptr [ebp-38], eax
'006cecfb    8d45c0                  lea eax, dword ptr [ebp-40]
'006cecfe    50                      push eax
'006cecff    c745d800000000          mov dword ptr [ebp-28], 00000000
'006ced06    c745c009000000          mov dword ptr [ebp-40], 00000009

' *** Reference to "rtcIsNull"
'006ced0d    ff1540114000            call dword ptr [00401140]
'006ced13    33ff                    xor edi, edi
var_num7 = Empty
'006ced15    668bf8                  mov di, ax
'006ced18    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006ced1b    f7d7                    not edi

' *** Reference to "__vbaFreeObj"
'006ced1d    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'006ced23    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeVar"
'006ced26    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_5)
'006ced2c    6685ff                  test edi, edi
'006ced2f    0f8410010000            je 6cee45

If (Not (IsNull(var_45))) Then
'006ced35    8b4508                  mov eax, dword ptr [ebp+08]
'006ced38    8b08                    mov ecx, dword ptr [eax]
'006ced3a    50                      push eax
'006ced3b    ff9108030000            call dword ptr [ecx+00000308]
'006ced41    50                      push eax
'006ced42    8d55d4                  lea edx, dword ptr [ebp-2c]
'006ced45    52                      push edx

' *** Reference to "__vbaObjSet"
'006ced46    ff15b4104000            call dword ptr [004010b4]
    Set var_44 = Me
'006ced4c    8d55dc                  lea edx, dword ptr [ebp-24]
'006ced4f    8985b4feffff            mov dword ptr [ebp+fffffeb4], eax
'006ced55    8b45e8                  mov eax, dword ptr [ebp-18]
'006ced58    8b08                    mov ecx, dword ptr [eax]
'006ced5a    52                      push edx
'006ced5b    50                      push eax
'006ced5c    ff91b4000000            call dword ptr [ecx+000000b4]
'006ced62    dbe2                    fnclex
'006ced64    85c0                    test ax, ax
'006ced66    7d11                    jge 6ced79
'006ced68    8b4de8                  mov ecx, dword ptr [ebp-18]
'006ced6b    68b4000000              push 000000b4
'006ced70    6830314300              push 00433130
'006ced75    51                      push ecx
'006ced76    50                      push eax
'006ced77    ffd6                    call esi
'006ced79    8b45dc                  mov eax, dword ptr [ebp-24]
'006ced7c    8b38                    mov edi, dword ptr [eax]
'006ced7e    8d5dd8                  lea ebx, dword ptr [ebp-28]
'006ced81    53                      push ebx
'006ced82    83ec10                  sub esp, 10
'006ced85    8bdc                    mov ebx, esp
'006ced87    ba08000000              mov edx, 00000008
'006ced8c    8913                    mov dword ptr [ebx], edx
'006ced8e    8b5594                  mov edx, dword ptr [ebp-6c]
'006ced91    895304                  mov dword ptr [ebx+04], edx
'006ced94    b944384500              mov ecx, 00453844
'006ced99    894b08                  mov dword ptr [ebx+08], ecx
'006ced9c    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006ced9f    50                      push eax
'006ceda0    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'006ceda6    894b0c                  mov dword ptr [ebx+0c], ecx
'006ceda9    ff5730                  call dword ptr [edi+30]
'006cedac    dbe2                    fnclex
'006cedae    85c0                    test ax, ax
'006cedb0    7d11                    jge 6cedc3
'006cedb2    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'006cedb8    6a30                    push 30
'006cedba    68d8304300              push 004330d8
'006cedbf    52                      push edx
'006cedc0    50                      push eax
'006cedc1    ffd6                    call esi
'006cedc3    8b45d8                  mov eax, dword ptr [ebp-28]
'006cedc6    8b08                    mov ecx, dword ptr [eax]
'006cedc8    8d55c0                  lea edx, dword ptr [ebp-40]
'006cedcb    52                      push edx
'006cedcc    50                      push eax
'006cedcd    8bf8                    mov edi, eax
'006cedcf    ff5144                  call dword ptr [ecx+44]
'006cedd2    dbe2                    fnclex
'006cedd4    85c0                    test ax, ax
'006cedd6    7d0b                    jge 6cede3
    
    If (    var_45) Then
'006cedd8    6a44                    push 44
'006cedda    6880304300              push 00433080
'006ceddf    57                      push edi
'006cede0    50                      push eax
'006cede1    ffd6                    call esi
    
End If
'006cede3    8bbdb4feffff            mov edi, dword ptr [ebp+fffffeb4]
'006cede9    8b1f                    mov ebx, dword ptr [edi]
'006cedeb    8d45c0                  lea eax, dword ptr [ebp-40]
'006cedee    50                      push eax

' *** Reference to "__vbaStrVarMove"
'006cedef    ff153c104000            call dword ptr [0040103c]
'006cedf5    8bd0                    mov edx, eax
'006cedf7    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006cedfa    ff15d0124000            call dword ptr [004012d0]
'006cee00    50                      push eax
'006cee01    57                      push edi
'006cee02    ff93a4000000            call dword ptr [ebx+000000a4]
'006cee08    dbe2                    fnclex
'006cee0a    85c0                    test ax, ax
'006cee0c    7d0e                    jge 6cee1c
'006cee0e    68a4000000              push 000000a4
'006cee13    68d00d4300              push 00430dd0
'006cee18    57                      push edi
'006cee19    50                      push eax
'006cee1a    ffd6                    call esi
'006cee1c    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006cee1f    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006cee25    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006cee28    51                      push ecx
'006cee29    8d55d8                  lea edx, dword ptr [ebp-28]
'006cee2c    52                      push edx
'006cee2d    8d45dc                  lea eax, dword ptr [ebp-24]
'006cee30    50                      push eax
'006cee31    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006cee33    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45, var_44)
'006cee39    83c410                  add esp, 10
'006cee3c    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeVar"
'006cee3f    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_5)
'ERROR: Two many next close:
End If
'006cee45    8b45e8                  mov eax, dword ptr [ebp-18]
'006cee48    8b08                    mov ecx, dword ptr [eax]
'006cee4a    8d55dc                  lea edx, dword ptr [ebp-24]
'006cee4d    52                      push edx
'006cee4e    50                      push eax
'006cee4f    ff91b4000000            call dword ptr [ecx+000000b4]
'006cee55    dbe2                    fnclex
'006cee57    85c0                    test ax, ax
'006cee59    7d11                    jge 6cee6c
'006cee5b    8b4de8                  mov ecx, dword ptr [ebp-18]
'006cee5e    68b4000000              push 000000b4
'006cee63    6830314300              push 00433130
'006cee68    51                      push ecx
'006cee69    50                      push eax
'006cee6a    ffd6                    call esi
'006cee6c    8b45dc                  mov eax, dword ptr [ebp-24]
'006cee6f    8b38                    mov edi, dword ptr [eax]
'006cee71    8d5dd8                  lea ebx, dword ptr [ebp-28]
'006cee74    53                      push ebx
'006cee75    83ec10                  sub esp, 10
'006cee78    8bdc                    mov ebx, esp
'006cee7a    ba08000000              mov edx, 00000008
'006cee7f    8913                    mov dword ptr [ebx], edx
'006cee81    8b5594                  mov edx, dword ptr [ebp-6c]
'006cee84    895304                  mov dword ptr [ebx+04], edx
'006cee87    b9a83b4500              mov ecx, 00453ba8
'006cee8c    894b08                  mov dword ptr [ebx+08], ecx
'006cee8f    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006cee92    50                      push eax
'006cee93    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'006cee99    894b0c                  mov dword ptr [ebx+0c], ecx
'006cee9c    ff5730                  call dword ptr [edi+30]
'006cee9f    dbe2                    fnclex
'006ceea1    85c0                    test ax, ax
'006ceea3    7d11                    jge 6ceeb6
'006ceea5    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'006ceeab    6a30                    push 30
'006ceead    68d8304300              push 004330d8
'006ceeb2    52                      push edx
'006ceeb3    50                      push eax
'006ceeb4    ffd6                    call esi
'006ceeb6    8b45d8                  mov eax, dword ptr [ebp-28]
'006ceeb9    8945c8                  mov dword ptr [ebp-38], eax
'006ceebc    8d45c0                  lea eax, dword ptr [ebp-40]
'006ceebf    50                      push eax
'006ceec0    c745d800000000          mov dword ptr [ebp-28], 00000000
'006ceec7    c745c009000000          mov dword ptr [ebp-40], 00000009

' *** Reference to "rtcIsNull"
'006ceece    ff1540114000            call dword ptr [00401140]
'006ceed4    33ff                    xor edi, edi
var_num7 = Empty
'006ceed6    668bf8                  mov di, ax
'006ceed9    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006ceedc    f7d7                    not edi

' *** Reference to "__vbaFreeObj"
'006ceede    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'006ceee4    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeVar"
'006ceee7    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_5)
'006ceeed    6685ff                  test edi, edi
'006ceef0    0f8410010000            je 6cf006

If (Not (IsNull(var_45))) Then
'006ceef6    8b4508                  mov eax, dword ptr [ebp+08]
'006ceef9    8b08                    mov ecx, dword ptr [eax]
'006ceefb    50                      push eax
'006ceefc    ff9104030000            call dword ptr [ecx+00000304]
'006cef02    50                      push eax
'006cef03    8d55d4                  lea edx, dword ptr [ebp-2c]
'006cef06    52                      push edx

' *** Reference to "__vbaObjSet"
'006cef07    ff15b4104000            call dword ptr [004010b4]
    Set var_44 = Me
'006cef0d    8d55dc                  lea edx, dword ptr [ebp-24]
'006cef10    8985b4feffff            mov dword ptr [ebp+fffffeb4], eax
'006cef16    8b45e8                  mov eax, dword ptr [ebp-18]
'006cef19    8b08                    mov ecx, dword ptr [eax]
'006cef1b    52                      push edx
'006cef1c    50                      push eax
'006cef1d    ff91b4000000            call dword ptr [ecx+000000b4]
'006cef23    dbe2                    fnclex
'006cef25    85c0                    test ax, ax
'006cef27    7d11                    jge 6cef3a
'006cef29    8b4de8                  mov ecx, dword ptr [ebp-18]
'006cef2c    68b4000000              push 000000b4
'006cef31    6830314300              push 00433130
'006cef36    51                      push ecx
'006cef37    50                      push eax
'006cef38    ffd6                    call esi
'006cef3a    8b45dc                  mov eax, dword ptr [ebp-24]
'006cef3d    8b38                    mov edi, dword ptr [eax]
'006cef3f    8d5dd8                  lea ebx, dword ptr [ebp-28]
'006cef42    53                      push ebx
'006cef43    83ec10                  sub esp, 10
'006cef46    8bdc                    mov ebx, esp
'006cef48    ba08000000              mov edx, 00000008
'006cef4d    8913                    mov dword ptr [ebx], edx
'006cef4f    8b5594                  mov edx, dword ptr [ebp-6c]
'006cef52    895304                  mov dword ptr [ebx+04], edx
'006cef55    b9a83b4500              mov ecx, 00453ba8
'006cef5a    894b08                  mov dword ptr [ebx+08], ecx
'006cef5d    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006cef60    50                      push eax
'006cef61    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'006cef67    894b0c                  mov dword ptr [ebx+0c], ecx
'006cef6a    ff5730                  call dword ptr [edi+30]
'006cef6d    dbe2                    fnclex
'006cef6f    85c0                    test ax, ax
'006cef71    7d11                    jge 6cef84
'006cef73    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'006cef79    6a30                    push 30
'006cef7b    68d8304300              push 004330d8
'006cef80    52                      push edx
'006cef81    50                      push eax
'006cef82    ffd6                    call esi
'006cef84    8b45d8                  mov eax, dword ptr [ebp-28]
'006cef87    8b08                    mov ecx, dword ptr [eax]
'006cef89    8d55c0                  lea edx, dword ptr [ebp-40]
'006cef8c    52                      push edx
'006cef8d    50                      push eax
'006cef8e    8bf8                    mov edi, eax
'006cef90    ff5144                  call dword ptr [ecx+44]
'006cef93    dbe2                    fnclex
'006cef95    85c0                    test ax, ax
'006cef97    7d0b                    jge 6cefa4
    
    If (    var_45) Then
'006cef99    6a44                    push 44
'006cef9b    6880304300              push 00433080
'006cefa0    57                      push edi
'006cefa1    50                      push eax
'006cefa2    ffd6                    call esi
    
End If
'006cefa4    8bbdb4feffff            mov edi, dword ptr [ebp+fffffeb4]
'006cefaa    8b1f                    mov ebx, dword ptr [edi]
'006cefac    8d45c0                  lea eax, dword ptr [ebp-40]
'006cefaf    50                      push eax

' *** Reference to "__vbaStrVarMove"
'006cefb0    ff153c104000            call dword ptr [0040103c]
'006cefb6    8bd0                    mov edx, eax
'006cefb8    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006cefbb    ff15d0124000            call dword ptr [004012d0]
'006cefc1    50                      push eax
'006cefc2    57                      push edi
'006cefc3    ff93a4000000            call dword ptr [ebx+000000a4]
'006cefc9    dbe2                    fnclex
'006cefcb    85c0                    test ax, ax
'006cefcd    7d0e                    jge 6cefdd
'006cefcf    68a4000000              push 000000a4
'006cefd4    68d00d4300              push 00430dd0
'006cefd9    57                      push edi
'006cefda    50                      push eax
'006cefdb    ffd6                    call esi
'006cefdd    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006cefe0    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006cefe6    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006cefe9    51                      push ecx
'006cefea    8d55d8                  lea edx, dword ptr [ebp-28]
'006cefed    52                      push edx
'006cefee    8d45dc                  lea eax, dword ptr [ebp-24]
'006ceff1    50                      push eax
'006ceff2    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006ceff4    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45, var_44)
'006ceffa    83c410                  add esp, 10
'006ceffd    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeVar"
'006cf000    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_5)
'ERROR: Two many next close:
End If
'006cf006    8b45e8                  mov eax, dword ptr [ebp-18]
'006cf009    8b08                    mov ecx, dword ptr [eax]
'006cf00b    8d55dc                  lea edx, dword ptr [ebp-24]
'006cf00e    52                      push edx
'006cf00f    50                      push eax
'006cf010    ff91b4000000            call dword ptr [ecx+000000b4]
'006cf016    dbe2                    fnclex
'006cf018    85c0                    test ax, ax
'006cf01a    7d11                    jge 6cf02d
'006cf01c    8b4de8                  mov ecx, dword ptr [ebp-18]
'006cf01f    68b4000000              push 000000b4
'006cf024    6830314300              push 00433130
'006cf029    51                      push ecx
'006cf02a    50                      push eax
'006cf02b    ffd6                    call esi
'006cf02d    8b45dc                  mov eax, dword ptr [ebp-24]
'006cf030    8b38                    mov edi, dword ptr [eax]
'006cf032    8d5dd8                  lea ebx, dword ptr [ebp-28]
'006cf035    53                      push ebx
'006cf036    83ec10                  sub esp, 10
'006cf039    8bdc                    mov ebx, esp
'006cf03b    ba08000000              mov edx, 00000008
'006cf040    8913                    mov dword ptr [ebx], edx
'006cf042    8b5594                  mov edx, dword ptr [ebp-6c]
'006cf045    895304                  mov dword ptr [ebx+04], edx
'006cf048    b9c03b4500              mov ecx, 00453bc0
'006cf04d    894b08                  mov dword ptr [ebx+08], ecx
'006cf050    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006cf053    50                      push eax
'006cf054    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'006cf05a    894b0c                  mov dword ptr [ebx+0c], ecx
'006cf05d    ff5730                  call dword ptr [edi+30]
'006cf060    dbe2                    fnclex
'006cf062    85c0                    test ax, ax
'006cf064    7d11                    jge 6cf077
'006cf066    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'006cf06c    6a30                    push 30
'006cf06e    68d8304300              push 004330d8
'006cf073    52                      push edx
'006cf074    50                      push eax
'006cf075    ffd6                    call esi
'006cf077    8b45d8                  mov eax, dword ptr [ebp-28]
'006cf07a    8945c8                  mov dword ptr [ebp-38], eax
'006cf07d    8d45c0                  lea eax, dword ptr [ebp-40]
'006cf080    50                      push eax
'006cf081    c745d800000000          mov dword ptr [ebp-28], 00000000
'006cf088    c745c009000000          mov dword ptr [ebp-40], 00000009

' *** Reference to "rtcIsNull"
'006cf08f    ff1540114000            call dword ptr [00401140]
'006cf095    33ff                    xor edi, edi
var_num7 = Empty
'006cf097    668bf8                  mov di, ax
'006cf09a    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006cf09d    f7d7                    not edi

' *** Reference to "__vbaFreeObj"
'006cf09f    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'006cf0a5    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeVar"
'006cf0a8    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_5)
'006cf0ae    6685ff                  test edi, edi
'006cf0b1    0f8410010000            je 6cf1c7

If (Not (IsNull(var_45))) Then
'006cf0b7    8b4508                  mov eax, dword ptr [ebp+08]
'006cf0ba    8b08                    mov ecx, dword ptr [eax]
'006cf0bc    50                      push eax
'006cf0bd    ff9100030000            call dword ptr [ecx+00000300]
'006cf0c3    50                      push eax
'006cf0c4    8d55d4                  lea edx, dword ptr [ebp-2c]
'006cf0c7    52                      push edx

' *** Reference to "__vbaObjSet"
'006cf0c8    ff15b4104000            call dword ptr [004010b4]
    Set var_44 = Me
'006cf0ce    8d55dc                  lea edx, dword ptr [ebp-24]
'006cf0d1    8985b4feffff            mov dword ptr [ebp+fffffeb4], eax
'006cf0d7    8b45e8                  mov eax, dword ptr [ebp-18]
'006cf0da    8b08                    mov ecx, dword ptr [eax]
'006cf0dc    52                      push edx
'006cf0dd    50                      push eax
'006cf0de    ff91b4000000            call dword ptr [ecx+000000b4]
'006cf0e4    dbe2                    fnclex
'006cf0e6    85c0                    test ax, ax
'006cf0e8    7d11                    jge 6cf0fb
'006cf0ea    8b4de8                  mov ecx, dword ptr [ebp-18]
'006cf0ed    68b4000000              push 000000b4
'006cf0f2    6830314300              push 00433130
'006cf0f7    51                      push ecx
'006cf0f8    50                      push eax
'006cf0f9    ffd6                    call esi
'006cf0fb    8b45dc                  mov eax, dword ptr [ebp-24]
'006cf0fe    8b38                    mov edi, dword ptr [eax]
'006cf100    8d5dd8                  lea ebx, dword ptr [ebp-28]
'006cf103    53                      push ebx
'006cf104    83ec10                  sub esp, 10
'006cf107    8bdc                    mov ebx, esp
'006cf109    ba08000000              mov edx, 00000008
'006cf10e    8913                    mov dword ptr [ebx], edx
'006cf110    8b5594                  mov edx, dword ptr [ebp-6c]
'006cf113    895304                  mov dword ptr [ebx+04], edx
'006cf116    b9c03b4500              mov ecx, 00453bc0
'006cf11b    894b08                  mov dword ptr [ebx+08], ecx
'006cf11e    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006cf121    50                      push eax
'006cf122    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'006cf128    894b0c                  mov dword ptr [ebx+0c], ecx
'006cf12b    ff5730                  call dword ptr [edi+30]
'006cf12e    dbe2                    fnclex
'006cf130    85c0                    test ax, ax
'006cf132    7d11                    jge 6cf145
'006cf134    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'006cf13a    6a30                    push 30
'006cf13c    68d8304300              push 004330d8
'006cf141    52                      push edx
'006cf142    50                      push eax
'006cf143    ffd6                    call esi
'006cf145    8b45d8                  mov eax, dword ptr [ebp-28]
'006cf148    8b08                    mov ecx, dword ptr [eax]
'006cf14a    8d55c0                  lea edx, dword ptr [ebp-40]
'006cf14d    52                      push edx
'006cf14e    50                      push eax
'006cf14f    8bf8                    mov edi, eax
'006cf151    ff5144                  call dword ptr [ecx+44]
'006cf154    dbe2                    fnclex
'006cf156    85c0                    test ax, ax
'006cf158    7d0b                    jge 6cf165
    
    If (    var_45) Then
'006cf15a    6a44                    push 44
'006cf15c    6880304300              push 00433080
'006cf161    57                      push edi
'006cf162    50                      push eax
'006cf163    ffd6                    call esi
    
End If
'006cf165    8bbdb4feffff            mov edi, dword ptr [ebp+fffffeb4]
'006cf16b    8b1f                    mov ebx, dword ptr [edi]
'006cf16d    8d45c0                  lea eax, dword ptr [ebp-40]
'006cf170    50                      push eax

' *** Reference to "__vbaStrVarMove"
'006cf171    ff153c104000            call dword ptr [0040103c]
'006cf177    8bd0                    mov edx, eax
'006cf179    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006cf17c    ff15d0124000            call dword ptr [004012d0]
'006cf182    50                      push eax
'006cf183    57                      push edi
'006cf184    ff93a4000000            call dword ptr [ebx+000000a4]
'006cf18a    dbe2                    fnclex
'006cf18c    85c0                    test ax, ax
'006cf18e    7d0e                    jge 6cf19e
'006cf190    68a4000000              push 000000a4
'006cf195    68d00d4300              push 00430dd0
'006cf19a    57                      push edi
'006cf19b    50                      push eax
'006cf19c    ffd6                    call esi
'006cf19e    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006cf1a1    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006cf1a7    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006cf1aa    51                      push ecx
'006cf1ab    8d55d8                  lea edx, dword ptr [ebp-28]
'006cf1ae    52                      push edx
'006cf1af    8d45dc                  lea eax, dword ptr [ebp-24]
'006cf1b2    50                      push eax
'006cf1b3    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006cf1b5    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45, var_44)
'006cf1bb    83c410                  add esp, 10
'006cf1be    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeVar"
'006cf1c1    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_5)
'ERROR: Two many next close:
End If
'006cf1c7    8b45e8                  mov eax, dword ptr [ebp-18]
'006cf1ca    8b08                    mov ecx, dword ptr [eax]
'006cf1cc    8d55dc                  lea edx, dword ptr [ebp-24]
'006cf1cf    52                      push edx
'006cf1d0    50                      push eax
'006cf1d1    ff91b4000000            call dword ptr [ecx+000000b4]
'006cf1d7    dbe2                    fnclex
'006cf1d9    85c0                    test ax, ax
'006cf1db    7d11                    jge 6cf1ee
'006cf1dd    8b4de8                  mov ecx, dword ptr [ebp-18]
'006cf1e0    68b4000000              push 000000b4
'006cf1e5    6830314300              push 00433130
'006cf1ea    51                      push ecx
'006cf1eb    50                      push eax
'006cf1ec    ffd6                    call esi
'006cf1ee    8b45dc                  mov eax, dword ptr [ebp-24]
'006cf1f1    8b38                    mov edi, dword ptr [eax]
'006cf1f3    8d5dd8                  lea ebx, dword ptr [ebp-28]
'006cf1f6    53                      push ebx
'006cf1f7    83ec10                  sub esp, 10
'006cf1fa    8bdc                    mov ebx, esp
'006cf1fc    ba08000000              mov edx, 00000008
'006cf201    8913                    mov dword ptr [ebx], edx
'006cf203    8b5594                  mov edx, dword ptr [ebp-6c]
'006cf206    895304                  mov dword ptr [ebx+04], edx
'006cf209    b9e03b4500              mov ecx, 00453be0
'006cf20e    894b08                  mov dword ptr [ebx+08], ecx
'006cf211    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006cf214    50                      push eax
'006cf215    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'006cf21b    894b0c                  mov dword ptr [ebx+0c], ecx
'006cf21e    ff5730                  call dword ptr [edi+30]
'006cf221    dbe2                    fnclex
'006cf223    85c0                    test ax, ax
'006cf225    7d11                    jge 6cf238
'006cf227    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'006cf22d    6a30                    push 30
'006cf22f    68d8304300              push 004330d8
'006cf234    52                      push edx
'006cf235    50                      push eax
'006cf236    ffd6                    call esi
'006cf238    8b45d8                  mov eax, dword ptr [ebp-28]
'006cf23b    8945c8                  mov dword ptr [ebp-38], eax
'006cf23e    8d45c0                  lea eax, dword ptr [ebp-40]
'006cf241    50                      push eax
'006cf242    c745d800000000          mov dword ptr [ebp-28], 00000000
'006cf249    c745c009000000          mov dword ptr [ebp-40], 00000009

' *** Reference to "rtcIsNull"
'006cf250    ff1540114000            call dword ptr [00401140]
'006cf256    33ff                    xor edi, edi
var_num7 = Empty
'006cf258    668bf8                  mov di, ax
'006cf25b    8d4ddc                  lea ecx, dword ptr [ebp-24]
'006cf25e    f7d7                    not edi

' *** Reference to "__vbaFreeObj"
'006cf260    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_39)
'006cf266    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeVar"
'006cf269    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_5)
'006cf26f    6685ff                  test edi, edi
'006cf272    0f8410010000            je 6cf388

If (Not (IsNull(var_45))) Then
'006cf278    8b4508                  mov eax, dword ptr [ebp+08]
'006cf27b    8b08                    mov ecx, dword ptr [eax]
'006cf27d    50                      push eax
'006cf27e    ff91fc020000            call dword ptr [ecx+000002fc]
'006cf284    50                      push eax
'006cf285    8d55d4                  lea edx, dword ptr [ebp-2c]
'006cf288    52                      push edx

' *** Reference to "__vbaObjSet"
'006cf289    ff15b4104000            call dword ptr [004010b4]
    Set var_44 = Me
'006cf28f    8d55dc                  lea edx, dword ptr [ebp-24]
'006cf292    8985b4feffff            mov dword ptr [ebp+fffffeb4], eax
'006cf298    8b45e8                  mov eax, dword ptr [ebp-18]
'006cf29b    8b08                    mov ecx, dword ptr [eax]
'006cf29d    52                      push edx
'006cf29e    50                      push eax
'006cf29f    ff91b4000000            call dword ptr [ecx+000000b4]
'006cf2a5    dbe2                    fnclex
'006cf2a7    85c0                    test ax, ax
'006cf2a9    7d11                    jge 6cf2bc
'006cf2ab    8b4de8                  mov ecx, dword ptr [ebp-18]
'006cf2ae    68b4000000              push 000000b4
'006cf2b3    6830314300              push 00433130
'006cf2b8    51                      push ecx
'006cf2b9    50                      push eax
'006cf2ba    ffd6                    call esi
'006cf2bc    8b45dc                  mov eax, dword ptr [ebp-24]
'006cf2bf    8b38                    mov edi, dword ptr [eax]
'006cf2c1    8d5dd8                  lea ebx, dword ptr [ebp-28]
'006cf2c4    53                      push ebx
'006cf2c5    83ec10                  sub esp, 10
'006cf2c8    8bdc                    mov ebx, esp
'006cf2ca    ba08000000              mov edx, 00000008
'006cf2cf    8913                    mov dword ptr [ebx], edx
'006cf2d1    8b5594                  mov edx, dword ptr [ebp-6c]
'006cf2d4    895304                  mov dword ptr [ebx+04], edx
'006cf2d7    b9e03b4500              mov ecx, 00453be0
'006cf2dc    894b08                  mov dword ptr [ebx+08], ecx
'006cf2df    8b4d9c                  mov ecx, dword ptr [ebp-64]
'006cf2e2    50                      push eax
'006cf2e3    8985c4feffff            mov dword ptr [ebp+fffffec4], eax
'006cf2e9    894b0c                  mov dword ptr [ebx+0c], ecx
'006cf2ec    ff5730                  call dword ptr [edi+30]
'006cf2ef    dbe2                    fnclex
'006cf2f1    85c0                    test ax, ax
'006cf2f3    7d11                    jge 6cf306
'006cf2f5    8b95c4feffff            mov edx, dword ptr [ebp+fffffec4]
'006cf2fb    6a30                    push 30
'006cf2fd    68d8304300              push 004330d8
'006cf302    52                      push edx
'006cf303    50                      push eax
'006cf304    ffd6                    call esi
'006cf306    8b45d8                  mov eax, dword ptr [ebp-28]
'006cf309    8b08                    mov ecx, dword ptr [eax]
'006cf30b    8d55c0                  lea edx, dword ptr [ebp-40]
'006cf30e    52                      push edx
'006cf30f    50                      push eax
'006cf310    8bf8                    mov edi, eax
'006cf312    ff5144                  call dword ptr [ecx+44]
'006cf315    dbe2                    fnclex
'006cf317    85c0                    test ax, ax
'006cf319    7d0b                    jge 6cf326
    
    If (    var_45) Then
'006cf31b    6a44                    push 44
'006cf31d    6880304300              push 00433080
'006cf322    57                      push edi
'006cf323    50                      push eax
'006cf324    ffd6                    call esi
    
End If
'006cf326    8bbdb4feffff            mov edi, dword ptr [ebp+fffffeb4]
'006cf32c    8b1f                    mov ebx, dword ptr [edi]
'006cf32e    8d45c0                  lea eax, dword ptr [ebp-40]
'006cf331    50                      push eax

' *** Reference to "__vbaStrVarMove"
'006cf332    ff153c104000            call dword ptr [0040103c]
'006cf338    8bd0                    mov edx, eax
'006cf33a    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaStrMove"
'006cf33d    ff15d0124000            call dword ptr [004012d0]
'006cf343    50                      push eax
'006cf344    57                      push edi
'006cf345    ff93a4000000            call dword ptr [ebx+000000a4]
'006cf34b    dbe2                    fnclex
'006cf34d    85c0                    test ax, ax
'006cf34f    7d0e                    jge 6cf35f
'006cf351    68a4000000              push 000000a4
'006cf356    68d00d4300              push 00430dd0
'006cf35b    57                      push edi
'006cf35c    50                      push eax
'006cf35d    ffd6                    call esi
'006cf35f    8d4de4                  lea ecx, dword ptr [ebp-1c]

' *** Reference to "__vbaFreeStr"
'006cf362    ff150c134000            call dword ptr [0040130c]
'#Cleanup(var_40)
'006cf368    8d4dd4                  lea ecx, dword ptr [ebp-2c]
'006cf36b    51                      push ecx
'006cf36c    8d55d8                  lea edx, dword ptr [ebp-28]
'006cf36f    52                      push edx
'006cf370    8d45dc                  lea eax, dword ptr [ebp-24]
'006cf373    50                      push eax
'006cf374    6a03                    push 03

' *** Reference to "__vbaFreeObjList"
'006cf376    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45, var_44)
'006cf37c    83c410                  add esp, 10
'006cf37f    8d4dc0                  lea ecx, dword ptr [ebp-40]

' *** Reference to "__vbaFreeVar"
'006cf382    ff1528104000            call dword ptr [00401028]
'#Cleanup(var_5)
'ERROR: Two many next close:
End If
'006cf388    8b45e8                  mov eax, dword ptr [ebp-18]
'006cf38b    8b08                    mov ecx, dword ptr [eax]
'006cf38d    50                      push eax
'006cf38e    ff91c4000000            call dword ptr [ecx+000000c4]
'006cf394    dbe2                    fnclex
'006cf396    85c0                    test ax, ax
'006cf398    7d11                    jge 6cf3ab
'006cf39a    8b55e8                  mov edx, dword ptr [ebp-18]
'006cf39d    68c4000000              push 000000c4
'006cf3a2    6830314300              push 00433130
'006cf3a7    52                      push edx
'006cf3a8    50                      push eax
'006cf3a9    ffd6                    call esi
'006cf3ab    c745fc00000000          mov dword ptr [ebp-04], 00000000
'006cf3b2    6803f46c00              push 006cf403
'006cf3b7    eb40                    jmp 6cf3f9
'006cf3b9    8d45e0                  lea eax, dword ptr [ebp-20]
'006cf3bc    50                      push eax
'006cf3bd    8d4de4                  lea ecx, dword ptr [ebp-1c]
'006cf3c0    51                      push ecx
'006cf3c1    6a02                    push 02

' *** Reference to "__vbaFreeStrList"
'006cf3c3    ff155c124000            call dword ptr [0040125c]
'#Cleanup( , -4536)
'006cf3c9    8d55d0                  lea edx, dword ptr [ebp-30]
'006cf3cc    52                      push edx
'006cf3cd    8d45d4                  lea eax, dword ptr [ebp-2c]
'006cf3d0    50                      push eax
'006cf3d1    8d4dd8                  lea ecx, dword ptr [ebp-28]
'006cf3d4    51                      push ecx
'006cf3d5    8d55dc                  lea edx, dword ptr [ebp-24]
'006cf3d8    52                      push edx
'006cf3d9    6a04                    push 04

' *** Reference to "__vbaFreeObjList"
'006cf3db    ff154c104000            call dword ptr [0040104c]
'#Cleanup( var_39, var_45, var_44, var_4)
'006cf3e1    8d45a0                  lea eax, dword ptr [ebp-60]
'006cf3e4    50                      push eax
'006cf3e5    8d4db0                  lea ecx, dword ptr [ebp-50]
'006cf3e8    51                      push ecx
'006cf3e9    8d55c0                  lea edx, dword ptr [ebp-40]
'006cf3ec    52                      push edx
'006cf3ed    6a03                    push 03

' *** Reference to "__vbaFreeVarList"
'006cf3ef    ff1540104000            call dword ptr [00401040]
'#Cleanup( var_5, var_6, var_7)
'006cf3f5    83c430                  add esp, 30
'006cf3f8    c3                      ret
'006cf3f9    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006cf3fc    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cf402    c3                      ret
'006cf403    8b4508                  mov eax, dword ptr [ebp+08]
'006cf406    8b08                    mov ecx, dword ptr [eax]
'006cf408    50                      push eax
'006cf409    ff5108                  call dword ptr [ecx+08]
'006cf40c    8b45fc                  mov eax, dword ptr [ebp-04]
'006cf40f    8b4dec                  mov ecx, dword ptr [ebp-14]
'006cf412    5f                      pop edi
'006cf413    5e                      pop esi
    'Reference to '__except_list'
'006cf414    64890d00000000          mov dword ptr fs:[00000000], ecx
'006cf41b    5b                      pop ebx
'006cf41c    8be5                    mov esp, ebp
'006cf41e    5d                      pop ebp
'006cf41f    c20400                  ret 0004


End Sub


'Event for BtnFermer
Private Sub BtnFermer_Click()
'006cddc0    55                      push ebp
'006cddc1    8bec                    mov ebp, esp
'006cddc3    83ec0c                  sub esp, 0c

' *** Reference to "__vbaExceptHandler"
'006cddc6    6866724000              push 00407266
'006cddcb    64a100000000            mov ax, word ptr fs:[00000000]
'006cddd1    50                      push eax
    'Reference to '__except_list'
'006cddd2    64892500000000          mov dword ptr fs:[00000000], esp
'006cddd9    83ec18                  sub esp, 18
'006cdddc    53                      push ebx
'006cdddd    56                      push esi
'006cddde    57                      push edi
'006cdddf    8965f4                  mov dword ptr [ebp-0c], esp
'006cdde2    c745f848684000          mov dword ptr [ebp-08], 00406848
'006cdde9    8b7d08                  mov edi, dword ptr [ebp+08]
'006cddec    8bc7                    mov eax, edi
'006cddee    83e001                  and eax, 01
'006cddf1    8945fc                  mov dword ptr [ebp-04], eax
'006cddf4    83e7fe                  and edi, fffffffe
'006cddf7    8b0f                    mov ecx, dword ptr [edi]
'006cddf9    57                      push edi
'006cddfa    897d08                  mov dword ptr [ebp+08], edi
'006cddfd    ff5104                  call dword ptr [ecx+04]
'006cde00    a124be7200              mov ax, word ptr [0072be24]
'006cde05    33db                    xor ebx, ebx
'006cde07    3bc3                    cmp eax, ebx
'006cde09    895de8                  mov dword ptr [ebp-18], ebx
'006cde0c    7510                    jne 6cde1e

If (0 = 0) Then
'006cde0e    6824be7200              push 0072be24
'006cde13    6870004300              push 00430070

' *** Reference to "__vbaNew2"
'006cde18    ff152c124000            call dword ptr [0040122c]
    Dim var_2 As New Global
End If
'006cde1e    8b3524be7200            mov esi, dword ptr [0072be24]
'006cde24    8b16                    mov edx, dword ptr [esi]
'006cde26    57                      push edi
'006cde27    8d45e8                  lea eax, dword ptr [ebp-18]
'006cde2a    50                      push eax
'006cde2b    8955d4                  mov dword ptr [ebp-2c], edx

' *** Reference to "__vbaObjSetAddref"
'006cde2e    ff15c8104000            call dword ptr [004010c8]
Set var_41 = Me
'006cde34    8b4dd4                  mov ecx, dword ptr [ebp-2c]
'006cde37    50                      push eax
'006cde38    56                      push esi
'006cde39    ff5110                  call dword ptr [ecx+10]
Call var_2.Unload(var_41)
'006cde3c    dbe2                    fnclex
'006cde3e    3bc3                    cmp eax, ebx
'006cde40    7d0f                    jge 6cde51
'006cde42    6a10                    push 10
'006cde44    6860004300              push 00430060
'006cde49    56                      push esi
'006cde4a    50                      push eax

' *** Reference to "__vbaHresultCheckObj"
'006cde4b    ff1580104000            call dword ptr [00401080]
'006cde51    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006cde54    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cde5a    895dfc                  mov dword ptr [ebp-04], ebx
'006cde5d    686fde6c00              push 006cde6f
'006cde62    eb0a                    jmp 6cde6e
'006cde64    8d4de8                  lea ecx, dword ptr [ebp-18]

' *** Reference to "__vbaFreeObj"
'006cde67    ff1508134000            call dword ptr [00401308]
'#Cleanup(var_41)
'006cde6d    c3                      ret
'006cde6e    c3                      ret
'006cde6f    8b4508                  mov eax, dword ptr [ebp+08]
'006cde72    8b10                    mov edx, dword ptr [eax]
'006cde74    50                      push eax
'006cde75    ff5208                  call dword ptr [edx+08]
'006cde78    8b45fc                  mov eax, dword ptr [ebp-04]
'006cde7b    8b4dec                  mov ecx, dword ptr [ebp-14]
'006cde7e    5f                      pop edi
'006cde7f    5e                      pop esi
    'Reference to '__except_list'
'006cde80    64890d00000000          mov dword ptr fs:[00000000], ecx
'006cde87    5b                      pop ebx
'006cde88    8be5                    mov esp, ebp
'006cde8a    5d                      pop ebp
'006cde8b    c20400                  ret 0004


End Sub


