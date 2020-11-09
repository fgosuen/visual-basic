VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ATPES01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   4980
   ClientLeft      =   4095
   ClientTop       =   4860
   ClientWidth     =   7290
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "ATPES01"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4980
   ScaleWidth      =   7290
   Begin VB.CommandButton Coleta 
      Caption         =   "Coleta"
      Height          =   315
      Left            =   3750
      TabIndex        =   29
      Top             =   210
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton CONTATOS 
      Caption         =   "Contatos"
      Height          =   315
      Left            =   5985
      TabIndex        =   31
      Top             =   210
      Width           =   900
   End
   Begin VB.CommandButton COMPLEMENTO 
      Caption         =   "Complemento"
      Height          =   315
      Left            =   4740
      TabIndex        =   30
      Top             =   210
      Width           =   1155
   End
   Begin VB.CommandButton BSel04 
      Caption         =   "?"
      Height          =   315
      Left            =   2370
      TabIndex        =   26
      Top             =   210
      Width           =   315
   End
   Begin VB.CommandButton NOVORAMO 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   35
      Top             =   3885
      Width           =   315
   End
   Begin VB.CommandButton NOVACIDADE 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      TabIndex        =   33
      Top             =   2385
      Width           =   315
   End
   Begin VB.CheckBox FLGMAL 
      Caption         =   "Etiqueta"
      Height          =   315
      Left            =   6000
      TabIndex        =   9
      Top             =   1665
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton NextPes 
      Caption         =   "+"
      Height          =   315
      Left            =   2700
      TabIndex        =   28
      Top             =   210
      Width           =   315
   End
   Begin VB.CommandButton BSel03 
      Caption         =   "?"
      Height          =   315
      Left            =   6240
      TabIndex        =   34
      Top             =   3885
      Width           =   315
   End
   Begin VB.TextBox DESRAMATV 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      MaxLength       =   30
      TabIndex        =   21
      Top             =   3885
      Width           =   4155
   End
   Begin VB.CommandButton Altera 
      Caption         =   "&Altera"
      Height          =   315
      Left            =   2535
      TabIndex        =   23
      Top             =   4380
      Width           =   945
   End
   Begin VB.CommandButton Cancela 
      Caption         =   "&Cancela"
      Height          =   315
      Left            =   5790
      TabIndex        =   25
      Top             =   4380
      Width           =   945
   End
   Begin VB.TextBox DESEND 
      Height          =   315
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   8
      Top             =   1665
      Width           =   4290
   End
   Begin VB.CommandButton Exclui 
      Caption         =   "&Exclui"
      Height          =   315
      Left            =   4155
      TabIndex        =   24
      Top             =   4380
      Width           =   945
   End
   Begin VB.CommandButton Inclui 
      Caption         =   "&Inclui"
      Height          =   315
      Left            =   915
      TabIndex        =   22
      Top             =   4380
      Width           =   945
   End
   Begin VB.TextBox INSEST 
      Height          =   315
      Left            =   4740
      MaxLength       =   20
      TabIndex        =   19
      Top             =   3465
      Width           =   2190
   End
   Begin VB.TextBox NOMBAI 
      Height          =   315
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   10
      Top             =   2025
      Width           =   2190
   End
   Begin VB.TextBox NOMCID 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2400
      MaxLength       =   25
      TabIndex        =   12
      Top             =   2385
      Width           =   2355
   End
   Begin VB.CommandButton BSel02 
      Caption         =   "?"
      Height          =   315
      Left            =   4800
      TabIndex        =   32
      Top             =   2385
      Width           =   315
   End
   Begin VB.TextBox NOMPES 
      Height          =   315
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   1
      Top             =   585
      Width           =   5355
   End
   Begin VB.TextBox NOMRES 
      Height          =   315
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   4
      Top             =   945
      Width           =   1665
   End
   Begin VB.TextBox NUMFAX 
      Height          =   315
      Left            =   4740
      MaxLength       =   20
      TabIndex        =   15
      Top             =   2745
      Width           =   2190
   End
   Begin VB.TextBox NUMIDT 
      Height          =   315
      Left            =   4740
      MaxLength       =   20
      TabIndex        =   17
      Top             =   3105
      Width           =   2190
   End
   Begin VB.TextBox NUMTLF 
      Height          =   315
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   14
      Top             =   2745
      Width           =   2190
   End
   Begin VB.TextBox RAZSOC 
      Height          =   315
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1305
      Width           =   5340
   End
   Begin VB.ComboBox TIPPES 
      Height          =   315
      ItemData        =   "ATPES01.frx":0000
      Left            =   3660
      List            =   "ATPES01.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   945
      Width           =   1350
   End
   Begin MSMask.MaskEdBox CODRAMATV 
      Height          =   315
      Left            =   1560
      TabIndex        =   20
      Top             =   3885
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   3
      Format          =   "000"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox CODCID 
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Top             =   2385
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   4
      Format          =   "0000"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox CODPES 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   210
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   6
      Format          =   "000000"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NUMCEP 
      Height          =   315
      Left            =   5910
      TabIndex        =   13
      Top             =   2385
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   9
      Format          =   "00000\.000"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NUMCGC 
      Height          =   315
      Left            =   1560
      TabIndex        =   18
      Top             =   3465
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   14
      Format          =   "00\.000\.000\/0000\-00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox DIAANV 
      Height          =   315
      Left            =   6060
      TabIndex        =   6
      Top             =   945
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   14
      Format          =   "00\/00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NUMCPF 
      Height          =   315
      Left            =   1560
      TabIndex        =   16
      Top             =   3105
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   11
      Format          =   "000\.000\.000\.00"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ramo de Ativid."
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   3945
      Width           =   1125
   End
   Begin VB.Label Lb015 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aniversário"
      Height          =   195
      Left            =   5220
      TabIndex        =   3
      Top             =   1005
      Width           =   780
   End
   Begin VB.Label Lb000 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      Height          =   195
      Left            =   360
      TabIndex        =   27
      Top             =   345
      Width           =   495
   End
   Begin VB.Label Lb001 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      Height          =   195
      Left            =   360
      TabIndex        =   36
      Top             =   705
      Width           =   420
   End
   Begin VB.Label Lb002 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      Height          =   195
      Left            =   3300
      TabIndex        =   37
      Top             =   1005
      Width           =   315
   End
   Begin VB.Label Lb003 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome Resumido"
      Height          =   195
      Left            =   360
      TabIndex        =   38
      Top             =   1005
      Width           =   1170
   End
   Begin VB.Label Lb004 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Razão Social"
      Height          =   195
      Left            =   360
      TabIndex        =   39
      Top             =   1365
      Width           =   945
   End
   Begin VB.Label Lb005 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço"
      Height          =   195
      Left            =   360
      TabIndex        =   40
      Top             =   1725
      Width           =   690
   End
   Begin VB.Label Lb006 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro"
      Height          =   195
      Left            =   360
      TabIndex        =   41
      Top             =   2085
      Width           =   405
   End
   Begin VB.Label Lb007 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade"
      Height          =   195
      Left            =   360
      TabIndex        =   42
      Top             =   2445
      Width           =   495
   End
   Begin VB.Label Lb008 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CEP"
      Height          =   195
      Left            =   5535
      TabIndex        =   43
      Top             =   2445
      Width           =   315
   End
   Begin VB.Label Lb009 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefone"
      Height          =   195
      Left            =   360
      TabIndex        =   44
      Top             =   2805
      Width           =   630
   End
   Begin VB.Label Lb010 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
      Height          =   195
      Left            =   4440
      TabIndex        =   45
      Top             =   2805
      Width           =   255
   End
   Begin VB.Label Lb011 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPF"
      Height          =   195
      Left            =   360
      TabIndex        =   46
      Top             =   3165
      Width           =   300
   End
   Begin VB.Label Lb012 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RG"
      Height          =   195
      Left            =   4440
      TabIndex        =   47
      Top             =   3165
      Width           =   240
   End
   Begin VB.Label Lb013 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CNPJ"
      Height          =   195
      Left            =   360
      TabIndex        =   48
      Top             =   3525
      Width           =   405
   End
   Begin VB.Label Lb014 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Insc.Estadual"
      Height          =   195
      Left            =   3735
      TabIndex        =   49
      Top             =   3525
      Width           =   960
   End
End
Attribute VB_Name = "ATPES01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim Entidade As String
Dim IncluiPessoa As Boolean
Dim IncluiTabela As Boolean
Dim codent As String
Dim noment As String
 
Dim TABPES As csRS
 
Dim TABCID As csRS
Dim TABRAT As csRS
Dim TABELA As csRS
 
Dim TABCLI As csRS
'Dim TABFOR As csRS
'Dim TABTRP As csRS
'Dim TABREP As csRS
 
Private Function PessoaOk() As Boolean
PessoaOk = False
If oIsEmpty(CODPES.Text) Then
   ShowMessage "Informe o código do Cliente ..."
 Else
   TABPES.RsSeek "=", CODPES.Text
   If TABPES.Nomatch Then
      ShowMessage "Cliente não cadastrado ..."
    Else
      PessoaOk = True
   End If
End If
End Function

Public Sub ShowForm(Tipo As String)
Dim s As String
HourGlassOn
Set TABPES = oTABPES()
Set TABCID = oTABCID()
Set TABRAT = oTABRAT()
 
Set TABCLI = oTABCLI()
'Set TABFOR = oTABFOR()
'Set TABTRP = oTABTRP()
'Set TABREP = oTABREP()
 
If Tipo = "C" Then
   Set TABELA = TABCLI
   s = "Clientes": codent = "codcli": noment = "nomcli"
 ElseIf Tipo = "F" Then
   'Set TABELA = TABFOR
   's = "Fornecedores": codent = "codfor": noment = "nomfor"
 ElseIf Tipo = "T" Then
   'Set TABELA = TABTRP
   'Coleta.Visible = True
   's = "Transportadoras": codent = "codtrp": noment = "nomtrp"
 ElseIf Tipo = "R" Then
   'Set TABELA = TABREP
   's = "Representantes": codent = "codrep": noment = "nomrep"
End If
 
Entidade = Tipo
 
Inclui.Enabled = False
Exclui.Enabled = False
Altera.Enabled = False
 
TIPPES.AddItem "Jurídica"
TIPPES.AddItem "Física"
TIPPES.ListIndex = 1
 
Me.Show
Me.Caption = ("Cadastro de " & s)
HourGlassOff
End Sub
 
Private Function CpfRepetido() As Boolean
Dim TABPES As csRS
 
Set TABPES = oTABPES(): TABPES.Index = "IDXPES04"
TABPES.RsSeek "=", NUMCPF.Text
If TABPES.Nomatch Then
   CpfRepetido = False
 Else
   If Val(CODPES.Text) = TABPES.RS!CODPES Then
      CpfRepetido = False
    Else
      If Val(NUMCPF.Text) <> 0 Then
         CpfRepetido = True
       Else
         CpfRepetido = False
      End If
   End If
End If
 
csRSClose TABPES
End Function
 
Private Function CgcRepetido() As Boolean
Dim TABPES As csRS
 
Set TABPES = oTABPES(): TABPES.Index = "IDXPES05"
TABPES.RsSeek "=", NUMCGC.Text
If TABPES.Nomatch Then
   CgcRepetido = False
 Else
   If Val(CODPES.Text) = TABPES.RS!CODPES Then
      CgcRepetido = False
    Else
      If Val(NUMCGC.Text) <> 0 Then
         CgcRepetido = True
       Else
         CgcRepetido = False
      End If
   End If
End If
 
csRSClose TABPES
End Function
 
Private Sub Altera_Click()
If Form_Ok() Then
   FormToDb
End If
End Sub
 
Private Sub BSel02_Click()
If BRCID01(TABCID.RS) Then
   CODCID.Text = TABCID.RS!CODCID
   NOMCID.Text = TABCID.RS!NOMCID
   NUMTLF.SetFocus
End If
End Sub
 
Private Sub BSel03_Click()
If BRRAT01(TABRAT.RS) Then
   CODRAMATV.Text = TABRAT.RS!CODRAMATV
   DESRAMATV.Text = TABRAT.RS!DESRAMATV
   CODRAMATV.SetFocus
End If
End Sub
 
Private Sub BSel04_Click()
Dim Ret As Boolean
 
Ret = False
If Entidade = "C" Then
   Ret = BRCLI01(TABELA.RS)
 ElseIf Entidade = "F" Then
   'Ret = BRFOR01(TABELA.RS)
 ElseIf Entidade = "T" Then
   'Ret = BRTRP01(TABELA.RS)
 ElseIf Entidade = "R" Then
   'Ret = BRREP01(TABELA.RS, MostraNivel:=False)
End If
 
If Ret Then
   CODPES.Text = TABELA.RS(codent)
   CODPES.SetFocus
   NOMPES.SetFocus
End If
End Sub
 
Private Sub Cancela_Click()
Unload Me
End Sub
 
Private Sub CODCID_KeyPress(KeyAscii As Integer)
VldEditNum KeyAscii, "####", (CODCID.Text)
End Sub
 
Private Sub CODCID_LostFocus()
TABCID.RsSeek "=", CODCID.Text
If Not TABCID.Nomatch Then
   NOMCID.Text = TABCID.RS!NOMCID
 Else
   NOMCID.Text = ""
End If
End Sub
 
Private Sub CODPES_KeyPress(KeyAscii As Integer)
VldEditNum KeyAscii, "######", (CODPES.Text)
End Sub
 
Private Sub CODPES_LostFocus()
Inclui.Enabled = False
Altera.Enabled = False
Exclui.Enabled = False
If Not oIsEmpty(CODPES.Text) Then
   TABPES.RsSeek "=", CODPES.Text
   If TABPES.Nomatch Then
      IncluiPessoa = True
      Inclui.Enabled = True
      IncluiTabela = True
    Else
      IncluiPessoa = False
      DbToForm
      TABELA.RsSeek "=", CODPES.Text
      If TABELA.Nomatch Then
         IncluiTabela = True
         Inclui.Enabled = True
       Else
         IncluiTabela = False
         Altera.Enabled = True
         Exclui.Enabled = True
      End If
   End If
End If
End Sub
 
Private Sub Cst_CODCID(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(CODCID.Text) Then
      ShowMessage "Informe o código da ciade..."
      Erro = True
      CODCID.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_CODPES(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(CODPES.Text) Then
      ShowMessage "Informe o código ..."
      Erro = True
      CODPES.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_INSEST(Erro As Boolean)
If Not Erro And INSEST.Enabled Then
   If oIsEmpty(INSEST.Text) Then
      ShowMessage "Inscricao estadual não informada ..."
      Erro = True
      INSEST.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_NOMPES(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(NOMPES.Text) Then
      ShowMessage "Nome não informado..."
      Erro = True
      NOMPES.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_NOMRES(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(NOMRES.Text) Then
      ShowMessage "Nome Resumido não informado ..."
      Erro = True
      NOMRES.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_NUMCEP(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(NUMCEP.Text) Then
      ShowMessage "Informe numero do CEP ..."
      Erro = True
      NUMCEP.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_DESEND(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(DESEND.Text) Then
      ShowMessage "Endereço não informado ..."
      Erro = True
      DESEND.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_NOMBAI(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(NOMBAI.Text) Then
      NOMBAI.Text = "-"
   End If
End If
End Sub
 
Private Sub Cst_NUMCGC(Erro As Boolean)
If Not Erro And NUMCGC.Enabled Then
   If oIsEmpty(NUMCGC.Text) Then
      ShowMessage "Informe o número do CNPJ..."
      Erro = True
      NUMCGC.SetFocus
    Else
      If Not CGCok(NUMCGC.Text) Then
         ShowMessage "Número do CNPJ inválido..."
         Erro = True
         NUMCGC.SetFocus
       ' Else
         ' If CgcRepetido() Then
         '   ShowMessage "CGC já existente..."
         '   Erro = True
         '    numcgc.SetFocus
         ' End If
      End If
   End If
End If
End Sub
 
Private Sub Cst_DIAANV(Erro As Boolean)
Dim D As Integer, m As Integer
If Not Erro And DIAANV.Enabled Then
   If Not oIsEmpty(DIAANV.Text) Then
      If Not ddmmOk(CInt(DIAANV.Text)) Then
         ShowMessage "Dia do aniversário inválido..."
         DIAANV.SetFocus
         Erro = True
      End If
   End If
End If
End Sub
 
Private Sub Cst_NUMCPF(Erro As Boolean)
If Not Erro And NUMCPF.Enabled Then
   If oIsEmpty(NUMCPF.Text) Then
      ShowMessage "Informe o número do CPF..."
      Erro = True
      NUMCPF.SetFocus
    Else
      If Not CPFok(NUMCPF.Text) Then
         ShowMessage "Número do CPF inválido..."
         Erro = True
         NUMCPF.SetFocus
       Else
         If CpfRepetido() Then
            ShowMessage "CPF já existente..."
            Erro = True
            NUMCPF.SetFocus
         End If
      End If
   End If
End If
End Sub
 
Private Sub Cst_NUMIDT(Erro As Boolean)
If Not Erro And NUMIDT.Enabled Then
   If oIsEmpty(NUMIDT.Text) Then
      NUMIDT.Text = "-"
   End If
End If
End Sub
 
Private Sub Cst_NUMTLF(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(NUMTLF.Text) Then
      NUMTLF.Text = "-"
   End If
End If
End Sub
 
Private Sub Cst_NUMFAX(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(NUMFAX.Text) Then
      NUMFAX.Text = "-"
   End If
End If
End Sub
 
Private Sub Cst_RAZSOC(Erro As Boolean)
If Not Erro And RAZSOC.Enabled Then
   If oIsEmpty(RAZSOC.Text) Then
      ShowMessage "Informe razao social ..."
      Erro = True
      RAZSOC.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_CODRAMATV(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(CODRAMATV.Text) Then
      ShowMessage "Informe o ramo de atividade ..."
      Erro = True
      CODRAMATV.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_Relacionamento(Erro As Boolean)
Dim cSql As String
 
If Not Erro Then
   TABCID.RsSeek "=", CODCID.Text
   If TABCID.Nomatch Then
      Erro = True
      ShowMessage "Cidade não cadastrada ..."
   End If
End If
 
If Not Erro Then
   TABRAT.RsSeek "=", CODRAMATV.Text
   If TABRAT.Nomatch Then
      Erro = True
      ShowMessage "Ramo de atividade não cadastrado ..."
   End If
End If
End Sub
 
Private Sub DbToForm()
CODPES.Text = TABPES.RS!CODPES
NOMPES.Text = TABPES.RS!NOMPES
TIPPES.ListIndex = TABPES.RS!TIPPES
NOMRES.Text = TABPES.RS!NOMRES
RAZSOC.Text = NullToSpc(TABPES.RS!RAZSOC)
DESEND.Text = TABPES.RS!DESEND
NOMBAI.Text = TABPES.RS!NOMBAI
CODCID.Text = TABPES.RS!CODCID
NUMCEP.Text = TABPES.RS!NUMCEP
NUMTLF.Text = NullToSpc(TABPES.RS!NUMTLF)
NUMFAX.Text = NullToSpc(TABPES.RS!NUMFAX)
NUMCPF.Text = NullToSpc(TABPES.RS!NUMCPF)
NUMIDT.Text = NullToSpc(TABPES.RS!NUMIDT)
NUMCGC.Text = NullToSpc(TABPES.RS!NUMCGC)
DIAANV.Text = NullToSpc(TABPES.RS!DIAANV)
INSEST.Text = NullToSpc(TABPES.RS!INSEST)
CODRAMATV.Text = TABPES.RS!CODRAMATV
FLGMAL.Value = IIf(TABPES.RS!FLGMAL, 1, 0)
 
TABCID.RsSeek "=", TABPES.RS!CODCID
If Not TABCID.Nomatch Then
   NOMCID.Text = TABCID.RS!NOMCID
 Else
   NOMCID.Text = ""
End If
 
TABRAT.RsSeek "=", TABPES.RS!CODRAMATV
If Not TABRAT.Nomatch Then
   DESRAMATV.Text = TABRAT.RS!DESRAMATV
 Else
   DESRAMATV.Text = ""
End If
 
End Sub
 
Private Sub CODRAMATV_KeyPress(KeyAscii As Integer)
VldEditNum KeyAscii, "###", (CODRAMATV.Text)
End Sub
 
Private Sub CODRAMATV_LostFocus()
TABRAT.RsSeek "=", CODRAMATV.Text
If Not TABRAT.Nomatch Then
   DESRAMATV.Text = TABRAT.RS!DESRAMATV
 Else
   DESRAMATV.Text = ""
End If
End Sub
 
Private Sub Coleta_Click()
If Not oIsEmpty(CODPES.Text) Then
   TABELA.RsSeek "=", Val(CODPES.Text)
   If Not TABELA.Nomatch Then
      Load ATECT01
      ATECT01.CODTRP.Text = NullToSpc(TABELA.RS!CODTRP)
      ATECT01.NOMTRP.Text = NullToSpc(TABELA.RS!NOMTRP)
      ATECT01.Show vbModal
    Else
      ShowMessage "Transportadora não Cadastrada ..."
      CODPES.SetFocus
   End If
Else
   ShowMessage "Informe o Código da Transportadora ..."
   CODPES.SetFocus
End If
End Sub
 
Private Sub COMPLEMENTO_Click()
If PessoaOk() Then
   ATDCM01.ShowForm (CODPES.Text)
End If
End Sub
 
Private Sub CONTATOS_Click()
If PessoaOk() Then
   Load ATPPC01
   ATPPC01.CODPES = CODPES.Text
   ATPPC01.NOMPES = TABPES.RS!NOMPES
   ATPPC01.NOMPESPCP = TABPES.RS!NOMPES
   ATPPC01.CODPES.Enabled = False
   ATPPC01.BSel01.Visible = False
   ATPPC01.Show vbModal
End If
End Sub
 
Private Sub Exclui_Click()
Dim TABPPC As csRS, TABDCM As csRS, t1 As New csTransaction
 
If OkToDelete() Then
   If oConfirme("Confirma exclusão ?") Then
      Set TABPPC = oTABPPC()
      Set TABDCM = oTABDCM()
 
      t1.BegTrans
 
      t1.Delete TABELA.RS
 
      If ExcluiPessoa() Then
         TABPPC.RsSeek ">=", TABPES.RS!CODPES
         If Not TABPPC.Nomatch Then
            Do While TABPPC.RS!CODPES = TABPES.RS!CODPES
               t1.Delete TABPPC.RS
               TABPPC.MoveNext
            Loop
         End If
 
         TABDCM.RsSeek "=", TABPES.RS!CODPES
         If Not TABDCM.Nomatch Then
            t1.Delete TABDCM.RS
         End If
 
         t1.Delete TABPES.RS
      End If
 
      t1.EndTrans
 
      If t1.ErrorCode = 0 Then
         Form_Clear
      End If
 
      csRSClose TABPPC
      csRSClose TABDCM
   End If
End If
End Sub
 
Private Sub Form_Clear()
CODPES.Text = ""
NOMPES.Text = ""
TIPPES.ListIndex = 0
NOMRES.Text = ""
RAZSOC.Text = ""
DESEND.Text = ""
NOMBAI.Text = ""
CODCID.Text = ""
NOMCID.Text = ""
NUMCEP.Text = ""
NUMTLF.Text = ""
NUMFAX.Text = ""
NUMCPF.Text = ""
NUMIDT.Text = ""
NUMCGC.Text = ""
DIAANV.Text = ""
INSEST.Text = ""
CODRAMATV.Text = ""
DESRAMATV.Text = ""
FLGMAL.Value = 1
CODPES.SetFocus
Inclui.Enabled = False
Exclui.Enabled = False
Altera.Enabled = False
End Sub
 
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   SendKeys Chr$(vbKeyTab)
End If
End Sub
 
Private Function Form_Ok() As Boolean
Dim Erro As Boolean
Erro = False
Cst_CODPES Erro
Cst_NOMPES Erro
Cst_NOMRES Erro
Cst_RAZSOC Erro
Cst_DESEND Erro
Cst_NOMBAI Erro
Cst_CODCID Erro
Cst_NUMCEP Erro
Cst_NUMTLF Erro
Cst_NUMFAX Erro
Cst_NUMCPF Erro
Cst_NUMIDT Erro
Cst_NUMCGC Erro
Cst_DIAANV Erro
Cst_INSEST Erro
Cst_CODRAMATV Erro
Cst_Relacionamento Erro
If Erro Then
   Form_Ok = False
 Else
   Form_Ok = True
End If
End Function
 
Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) \ 2
Me.Left = (Screen.Width - Me.Width) \ 2
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
csRSClose TABPES
csRSClose TABCID
csRSClose TABCLI
csRSClose TABRAT
End Sub
 
Private Sub FormToDb()
Dim t1 As New csTransaction, Ret As Boolean
 
HourGlassOn
 
t1.BegTrans
 
If IncluiPessoa Then
   Ret = t1.AddNew(TABPES.RS)
 Else
   Ret = t1.Edit(TABPES.RS)
End If
If Ret Then
   TABPES.RS!CODPES = CODPES.Text
   TABPES.RS!NOMPES = NOMPES.Text
   TABPES.RS!TIPPES = TIPPES.ListIndex
   TABPES.RS!NOMRES = NOMRES.Text
   TABPES.RS!RAZSOC = SpcToNull(RAZSOC.Text)
   TABPES.RS!DESEND = DESEND.Text
   TABPES.RS!NOMBAI = NOMBAI.Text
   TABPES.RS!CODCID = CODCID.Text
   TABPES.RS!NUMCEP = NUMCEP.Text
   TABPES.RS!NUMTLF = NUMTLF.Text
   TABPES.RS!NUMFAX = NUMFAX.Text
   TABPES.RS!NUMCPF = SpcToNull(NUMCPF.Text)
   TABPES.RS!NUMIDT = SpcToNull(NUMIDT.Text)
   TABPES.RS!NUMCGC = SpcToNull(NUMCGC.Text)
   TABPES.RS!DIAANV = SpcToNull(DIAANV.Text)
   TABPES.RS!INSEST = SpcToNull(INSEST.Text)
   TABPES.RS!CODRAMATV = CODRAMATV.Text
   TABPES.RS!FLGMAL = IIf(FLGMAL.Value = 1, True, False)
   t1.Update TABPES.RS
 
   If IncluiTabela Then
      Ret = t1.AddNew(TABELA.RS)
    Else
      Ret = t1.Edit(TABELA.RS)
   End If
 
   If Ret Then
      TABCLI.RS!CODCLI = CODPES.Text
      TABCLI.RS!NOMCLI = NOMPES.Text
      t1.Update TABELA.RS
 
      Call Atualiza(TABCLI, t1, "nomcli")
      'Atualiza TABFOR, T1, "nomfor"
      'Atualiza TABTRP, T1, "nomtrp"
      'Atualiza TABREP, T1, "nomrep"
   End If
End If
 
t1.EndTrans
 
If t1.ErrorCode = 0 Then
    Form_Clear
End If
 
HourGlassOff
End Sub
 
Private Sub Atualiza(oRS As csRS, t1 As csTransaction, noment As String)
TABCLI.RsSeek "=", CODPES.Text
If Not TABCLI.Nomatch Then
   If t1.Edit(TABCLI.RS) Then
      TABCLI.RS!NOMCLI = NOMPES.Text
      t1.Update TABCLI.RS
   End If
End If
End Sub
 
Private Sub Inclui_Click()
If Form_Ok() Then
   FormToDb
End If
End Sub
 
Private Sub NextPes_Click()
On Error GoTo FimNextPes
TABPES.MoveLast
If TABPES.NoCurrentRecord Then
   CODPES.Text = "1"
 Else
   CODPES = TABPES.RS!CODPES + 1
End If
CODPES.SetFocus
NOMPES.SetFocus
Exit Sub
 
FimNextPes:
CODPES.Text = "1"
CODPES.SetFocus
End Sub
 
Private Sub NOVACIDADE_Click()
ATCID01.Show
End Sub
 
Private Sub NOVORAMO_Click()
ATRAT01.Show
End Sub
 
Private Sub NUMCEP_KeyPress(KeyAscii As Integer)
VldEditNum KeyAscii, "########", (NUMCEP.Text)
End Sub
 
Private Sub NUMCGC_KeyPress(KeyAscii As Integer)
VldEditNum KeyAscii, "##############", (NUMCGC.Text)
End Sub
 
Private Sub DIAANV_KeyPress(KeyAscii As Integer)
VldEditNum KeyAscii, "####", (DIAANV.Text)
End Sub
 
Private Sub NUMCPF_KeyPress(KeyAscii As Integer)
VldEditNum KeyAscii, "###########", (NUMCPF.Text)
End Sub
 
Private Function ExcluiPessoa() As Boolean
Dim FlagOk As Boolean
 
FlagOk = True
 
If FlagOk Then
   TABCLI.RsSeek "=", TABPES.RS!CODPES
   If Not TABCLI.Nomatch Then
      FlagOk = False
   End If
End If
 
'If FlagOk Then
'   TABFOR.RsSeek "=", TABPES.RS!CODPES
'   If Not TABFOR.Nomatch Then
'      FlagOk = False
'   End If
'End If
 
'If FlagOk Then
'   TABTRP.RsSeek "=", TABPES.RS!CODPES
'   If Not TABTRP.Nomatch Then
'      FlagOk = False
'   End If
'End If
 
'If FlagOk Then
'   TABREP.RsSeek "=", TABPES.RS!CODPES
'   If Not TABREP.Nomatch Then
'      FlagOk = False
'   End If
'End If
 
ExcluiPessoa = FlagOk
End Function
 
Private Function OkToDelete() As Boolean
OkToDelete = True
End Function
 
 
Private Sub TIPPES_Click()
If TIPPES.ListIndex = 0 Then
   NUMCPF.Enabled = False
   NUMIDT.Enabled = False
   DIAANV.Enabled = False
   RAZSOC.Enabled = True
   NUMCGC.Enabled = True
   INSEST.Enabled = True
   Lb011.Enabled = False
   Lb012.Enabled = False
   Lb013.Enabled = True
   Lb014.Enabled = True
   Lb015.Enabled = False
   Lb004.Enabled = True
   NUMCPF.Text = ""
   NUMIDT.Text = ""
   DIAANV.Text = ""
 Else
   NUMCPF.Enabled = True
   NUMIDT.Enabled = True
   DIAANV.Enabled = True
   RAZSOC.Enabled = False
   NUMCGC.Enabled = False
   INSEST.Enabled = True
   Lb011.Enabled = True
   Lb012.Enabled = True
   Lb013.Enabled = False
   Lb014.Enabled = True
   Lb015.Enabled = True
   Lb004.Enabled = False
   NUMCGC.Text = ""
   INSEST.Text = ""
   RAZSOC.Text = ""
End If
End Sub
