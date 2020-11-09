VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form RPCLI01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Clientes"
   ClientHeight    =   3330
   ClientLeft      =   3600
   ClientTop       =   2850
   ClientWidth     =   7110
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3330
   ScaleWidth      =   7110
   Begin VB.Frame TIPREL 
      Caption         =   " Tipo do &Relatório "
      Height          =   1320
      Left            =   360
      TabIndex        =   7
      Top             =   1035
      Width           =   2040
      Begin VB.OptionButton TIPCOM 
         Caption         =   "Completo"
         Height          =   240
         Left            =   270
         TabIndex        =   10
         Top             =   900
         Width           =   1365
      End
      Begin VB.OptionButton TIPFIS 
         Caption         =   "Dados Fiscais"
         Height          =   195
         Left            =   270
         TabIndex        =   9
         Top             =   630
         Width           =   1365
      End
      Begin VB.OptionButton TIPEND 
         Caption         =   "Endereço"
         Height          =   240
         Left            =   270
         TabIndex        =   8
         Top             =   315
         Value           =   -1  'True
         Width           =   1365
      End
   End
   Begin VB.CommandButton cancelar 
      Caption         =   "&Cancela"
      Height          =   390
      Left            =   5055
      TabIndex        =   19
      Top             =   2655
      Width           =   1140
   End
   Begin VB.CommandButton imprimir 
      Caption         =   "&Imprime"
      Height          =   390
      Left            =   3045
      TabIndex        =   18
      Top             =   2655
      Width           =   1140
   End
   Begin VB.TextBox NOMCLI 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2115
      TabIndex        =   2
      Top             =   225
      Width           =   4335
   End
   Begin VB.CommandButton ConsultaCODCLI 
      Caption         =   "?"
      Height          =   300
      Left            =   6525
      TabIndex        =   20
      Top             =   225
      Width           =   330
   End
   Begin VB.Frame ORDREG 
      Caption         =   " &Ordem "
      Height          =   1320
      Left            =   4770
      TabIndex        =   15
      Top             =   1035
      Width           =   1725
      Begin VB.OptionButton CODIGO 
         Caption         =   "Código"
         Height          =   240
         Left            =   270
         TabIndex        =   17
         Top             =   675
         Width           =   1095
      End
      Begin VB.OptionButton ALFABETICA 
         Caption         =   "Alfabética"
         Height          =   240
         Left            =   270
         TabIndex        =   16
         Top             =   405
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox INTINI 
      Height          =   285
      Left            =   1125
      MaxLength       =   20
      TabIndex        =   4
      Text            =   "A"
      Top             =   585
      Width           =   2355
   End
   Begin VB.TextBox INTFIN 
      Height          =   285
      Left            =   3960
      MaxLength       =   20
      TabIndex        =   6
      Text            =   "Z"
      Top             =   585
      Width           =   2490
   End
   Begin VB.Frame TIPPES 
      Caption         =   " &Tipo de Cliente "
      Height          =   1320
      Left            =   2610
      TabIndex        =   11
      Top             =   1035
      Width           =   1995
      Begin VB.OptionButton CLIFIS 
         Caption         =   "Pessoa Física"
         Height          =   195
         Left            =   270
         TabIndex        =   13
         Top             =   620
         Width           =   1545
      End
      Begin VB.OptionButton CLIJUR 
         Caption         =   "Pessoa Jurídica"
         Height          =   195
         Left            =   270
         TabIndex        =   14
         Top             =   900
         Width           =   1545
      End
      Begin VB.OptionButton CLIAMB 
         Caption         =   "Ambos"
         Height          =   240
         Left            =   270
         TabIndex        =   12
         Top             =   315
         Value           =   -1  'True
         Width           =   1545
      End
   End
   Begin MSMask.MaskEdBox CODCLI 
      Height          =   300
      Left            =   1125
      TabIndex        =   1
      Top             =   225
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   529
      _Version        =   393216
      Format          =   "000000"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Inter&valo"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   630
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Clie&nte"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   270
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "&até"
      Height          =   195
      Left            =   3600
      TabIndex        =   5
      Top             =   630
      Width           =   225
   End
End
Attribute VB_Name = "RPCLI01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim TABPES As csRS, TABRAT As csRS
Dim TABCLI As csRS, TABCID As csRS
Dim sTitle As String
 
Private Sub DisableItens()
   TIPPES.Enabled = False
   ORDREG.Enabled = False
   CLIFIS.Enabled = False
   CLIJUR.Enabled = False
   CLIAMB.Enabled = False
   ALFABETICA.Enabled = False
   CODIGO.Enabled = False
   Label1.Enabled = False
   Label3.Enabled = False
   INTINI.Enabled = False
   INTFIN.Enabled = False
End Sub
 
Private Sub EnableItens()
   TIPPES.Enabled = True
   ORDREG.Enabled = True
   CLIFIS.Enabled = True
   CLIJUR.Enabled = True
   CLIAMB.Enabled = True
   ALFABETICA.Enabled = True
   CODIGO.Enabled = True
   Label1.Enabled = True
   Label3.Enabled = True
   INTINI.Enabled = True
   INTFIN.Enabled = True
End Sub
Private Sub MontaSql()
Dim sIntIni As String, sIntFin As String, sSql As String
Dim TABREP As csRS
 
sIntIni = Chr$(34) & INTINI.Text & Chr$(34)
sIntFin = Chr$(34) & INTFIN.Text & Chr$(255) & Chr$(34)
 
If Not oIsEmpty(Me.NOMCLI.Text) Then
   sSql = "SELECT TABCLI.*, TABPES.* " & _
          "FROM TABCLI LEFT JOIN TABPES ON TABCLI.CODCLI = TABPES.CODPES " & _
          "WHERE NOMCLI = " & Chr$(34) & Me.NOMCLI.Text & Chr$(34)
Else
   sSql = "SELECT TABCLI.*, TABPES.* " & _
          "FROM TABCLI LEFT JOIN TABPES ON TABCLI.CODCLI = TABPES.CODPES " & _
          "WHERE " & _
             "NOMCLI >= " & sIntIni & " AND " & _
             "NOMCLI <= " & sIntFin & "  "
 
   If CLIJUR.Value Then
      sSql = sSql & "AND TIPPES = 0 "
   ElseIf CLIFIS.Value Then
      sSql = sSql & "AND TIPPES = 1 "
   End If
 
   If ALFABETICA.Value Then
      sSql = sSql & "ORDER BY NOMCLI"
   Else
      sSql = sSql & "ORDER BY CODCLI"
   End If
 
End If
 
Set TABCLI = oTABCLI(sSql)
 
End Sub
Private Sub PrintCabecalho(oPrinter As csPrinter, nPag As Integer)
nPag = nPag + 1
If nPag <> 1 Then
   oPrinter.NewPage
End If
With oPrinter
   .CurLine = 0
   .DrawLineXL 0, .LineWidth
 
   .CurLine = 1
   .PrintCenterXL G_RL_Empresa
   .LPrintXL 0, Format$(Now, "dd/mm/yyyy HH:MM")
   .LPrintXL -.LineWidth, "PAG. " & Format$(nPag, "00000")
 
   .CurLine = 2
   .PrintCenterXL .ReportTitle
 
   .CurLine = 3
   .DrawLineXL 0, .LineWidth
   .CurLine = 5
End With
End Sub
Private Sub PrintDetalhe(oPrinter As csPrinter)
   With oPrinter
      'Impressão das linhas de detalhe ...
      .LPrintXL 0, "Cliente: " & Format$(TABCLI.RS!CODCLI, "000000") & " - " & NullToSpc(TABCLI.RS!NOMCLI)
      .LPrintXL 100, "Limite Crédito: " & FormatMoney(TABCLI.RS!LIMCRE)
 
      .CurLine = .CurLine + 1
 
      .LPrintXL 0, "Razão Social: " & NullToSpc(TABCLI.RS!RAZSOC)
      .LPrintXL 100, "Nome Resumido: " & NullToSpc(TABCLI.RS!NOMRES)
      If TABCLI.RS!TIPPES = 1 Then
         .LPrintXL 165, "Aniversário: " & NullToSpc(TABCLI.RS!DIAANV)
      End If
 
      .CurLine = .CurLine + 1
      .LPrintXL 0, "Endereço: " & RTrim(NullToSpc(TABCLI.RS!DESEND))
      .LPrintXL 100, "Bairro: " & IIf(IsNull(TABCLI.RS!NOMBAI), "", NullToSpc(TABCLI.RS!NOMBAI))
      .LPrintXL 165, "CEP: " & Format$(NullToZero(TABCLI.RS!NUMCEP), "00000\-000")
 
      .CurLine = .CurLine + 1
      .LPrintXL 0, "Cidade: " & RetornaNOMCID(TABCID, TABCLI.RS!CODCID) & " - " & RetornaUNDFED(TABCID, TABCLI.RS!CODCID)
      .LPrintXL 100, "Telefone: " & RTrim(NullToSpc(TABCLI.RS!NUMTLF))
      .LPrintXL 165, "Fax: " & RTrim(NullToSpc(TABCLI.RS!NUMFAX))
 
      .CurLine = .CurLine + 1
      .LPrintXL 0, "Tipo: " & IIf(TABCLI.RS!TIPPES = 0, "Pessoa Jurídica", "Pessoa Física")
      If TABCLI.RS!TIPPES = 1 Then
         .LPrintXL 50, "CPF: " & Format$(NullToZero(TABCLI.RS!NUMCPF), "000\.000\.000\-00")
         .LPrintXL 100, "Identidade: " & NullToSpc(TABCLI.RS!NUMIDT)
         .LPrintXL 165, "Aniversário: " & Format$(NullToZero(TABCLI.RS!DIAANV), "00/00")
      Else
         .LPrintXL 50, "CNPJ: " & Format$(TABCLI.RS!NUMCGC, "00\.000\.000\/0000\-00")
         .LPrintXL 100, "Inscrição Estadual: " & NullToSpc(TABCLI.RS!INSEST)
      End If
 
      .CurLine = .CurLine + 1
      .LPrintXL 0, "Ramo Atividade: " & RetornaDESRAMATV(TABRAT, TABCLI.RS!CODRAMATV)
 
      .CurLine = .CurLine + 2
   End With
End Sub
Private Sub cancelar_Click()
   Unload Me
End Sub
Private Sub CODCLI_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF9 Then
      ConsultaCODCLI_Click
   End If
End Sub
Private Sub CODCLI_KeyPress(KeyAscii As Integer)
   VldEditNum KeyAscii, "#####", (CODCLI.Text)
End Sub
Private Sub CODCLI_LostFocus()
TABCLI.RsSeek "=", Me.CODCLI.Text
If TABCLI.Nomatch Then
   Me.NOMCLI.Text = ""
   EnableItens
Else
   Me.NOMCLI.Text = NullToSpc(TABCLI.RS!NOMCLI)
   DisableItens
End If
End Sub
Private Sub ConsultaCODCLI_Click()
If BRCLI01(TABCLI.RS) Then
   Me.CODCLI.Text = TABCLI.RS!CODCLI
   Me.NOMCLI.Text = TABCLI.RS!NOMCLI
End If
Me.CODCLI.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   SendKeys Chr$(vbKeyTab)
End If
End Sub

Private Sub Form_Load()
HourGlassOn
Set TABCLI = oTABCLI()
Set TABCID = oTABCID()
Set TABRAT = oTABRAT()

Me.Top = (Screen.Height - Me.Height) \ 2
Me.Left = (Screen.Width - Me.Width) \ 2

HourGlassOff
End Sub

Private Sub Form_Unload(Cancel As Integer)
csRSClose TABCLI
End Sub

Private Sub imprimir_Click()
If TIPEND.Value Then
   PrintRelEndereco
ElseIf TIPFIS.Value Then
   PrintRelFiscal
Else
   PrintRelCompleto
End If
End Sub

Private Sub PrintRelEndereco()
Dim oReport As New csReport

With oReport
   .Prefix = AppSpoolPrefix
   .Path = "."
   .ReportName = "RPCLI01"
   .ReportTitle2 = "Cadastro de Clientes - Endereço"
   .ReportTitle1 = G_RL_Empresa
   .FinalBreak = True
   If .ShowDialogSets() Then
      HourGlassOn
      Me.Visible = False
      MontaSql
      If TABCLI.EOF Then
         iMsgBox "Nenhum registro processado ..."
       Else
         .AddColumn "Código", 8, "000000", csRight, False
         .AddColumn "Nome", 40, "", csLeft, False
         .AddColumn "Endereço", 40, "", csLeft, False
         .AddColumn "Bairro", 20, "", csLeft, False
         .AddColumn "Cidade", 25, "", csLeft, False
         .AddColumn "UF", 3, "", csLeft, False
         .AddColumn "CEP", 11, "00000\-000", csLeft, False
         .AddColumn "Telefone", 20, "", csLeft, False
         Do While Not TABCLI.EOF
            .ColumnValue 1, TABCLI.RS!CODCLI, vbBlack
            .ColumnValue 2, TABCLI.RS!NOMCLI, vbBlack
            .ColumnValue 3, TABCLI.RS!DESEND, vbBlack
            .ColumnValue 4, TABCLI.RS!NOMBAI, vbBlack
            .ColumnValue 5, RetornaNOMCID(TABCID, TABCLI.RS!CODCID), vbBlack
            .ColumnValue 6, RetornaUNDFED(TABCID, TABCLI.RS!CODCID), vbBlack
            .ColumnValue 7, TABCLI.RS!NUMCEP, vbBlack
            .ColumnValue 8, TABCLI.RS!NUMTLF, vbBlack
            .PrintDetail
            TABCLI.MoveNext
         Loop
         .EndReport
      End If
      HourGlassOff
      Unload Me
   End If
End With
 
Set oReport = Nothing
End Sub

Private Sub PrintRelFiscal()
Dim oReport As New csReport, sTitle2 As String
With oReport
   sTitle = "Cadastro de Clientes x Dados Fiscais - "
   .ReportName = "RPCLI02"
   .ReportTitle1 = G_RL_Empresa
   .ReportTitle2 = "Cadastro de Clientes"
   If .ShowDialogSets() Then
      HourGlassOn
      MontaSql
      If CLIFIS.Value Then
         .ReportTitle2 = sTitle & "Pessoa Física"
         sTitle2 = "CPF"
      ElseIf CLIJUR.Value Then
         .ReportTitle2 = sTitle & "Pessoa Jurídica"
         sTitle2 = "CNPJ"
      Else
         .ReportTitle2 = sTitle & "Ambos"
         sTitle2 = "CNPJ/CPF"
      End If
 
      Me.Visible = False
      If TABCLI.EOF Then
         Beep
         MsgBox "Nenhum Registro Processado .RS!", vbInformation, "Informação"
      Else
         .Prefix = AppSpoolPrefix
         .Path = "."
         .AddColumn "Código", 7, "000000"
         .AddColumn "Cliente", 40
         .AddColumn sTitle2, 22
         .AddColumn "Inscrição Estadual", 20
         .AddColumn "Identidade", 20
         .AddColumn "Telefone", 20
 
         While Not TABCLI.EOF
            .ColumnValue 1, TABCLI.RS!CODCLI, vbBlack
            .ColumnValue 2, TABCLI.RS!NOMCLI, vbBlue
            If TABCLI.RS!TIPPES = 1 Then
               .ColumnValue 3, Format$(TABCLI.RS!NUMCPF, "000\.000\.000\-00"), vbBlack
            Else
               .ColumnValue 3, Format$(TABCLI.RS!NUMCGC, "00\.000\.000\/0000\-00"), vbBlack
            End If
            .ColumnValue 4, TABCLI.RS!INSEST, vbBlack
            .ColumnValue 5, TABCLI.RS!NUMIDT, vbBlack
            .ColumnValue 6, TABCLI.RS!NUMTLF, vbBlue
            .PrintDetail
            TABCLI.MoveNext
         Wend
         .EndReport
      End If
      Unload Me
   End If
End With
HourGlassOff
End Sub

Private Sub PrintRelCompleto()
Dim oPrinter As New csPrinter
Dim nPag As Integer
oPrinter.ReportName = "RPCLI03"
oPrinter.ReportTitle = "Cadastro de Clientes - Completo"
If oPrinter.ShowDialogSets() Then
   HourGlassOn
   Me.Visible = False
   MontaSql
   Me.Visible = False
   If TABCLI.EOF Then
      Beep
      MsgBox "Nenhum Registro Processado .RS!", vbInformation, "Informação"
   Else
      With oPrinter
         .Prefix = AppSpoolPrefix
         .Path = "."
         .PrintOpen
         nPag = 0
         .CurLine = 500
         While Not TABCLI.EOF
            If .CurLine + 8 > .Lines Then
               PrintCabecalho oPrinter, nPag
            End If
            PrintDetalhe oPrinter
            TABCLI.MoveNext
         Wend
         .PrintClose
      End With
   End If
   Unload Me
End If
Set oPrinter = Nothing
HourGlassOff
End Sub

Private Sub INTFIN_LostFocus()
   If oIsEmpty(Me.INTFIN.Text) Then
      Me.INTFIN.Text = "Z"
   End If
End Sub

Private Sub INTINI_LostFocus()
   If oIsEmpty(Me.INTINI.Text) Then
      Me.INTINI.Text = "A"
   End If
End Sub
