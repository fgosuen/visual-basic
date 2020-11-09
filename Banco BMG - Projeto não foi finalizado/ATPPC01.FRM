VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ATPPC01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pessoas de Contato"
   ClientHeight    =   2910
   ClientLeft      =   2355
   ClientTop       =   2475
   ClientWidth     =   7020
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "ATPPC01"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2910
   ScaleWidth      =   7020
   Begin VB.CommandButton Altera 
      Caption         =   "Altera"
      Height          =   315
      Left            =   5280
      TabIndex        =   8
      Top             =   1560
      Width           =   945
   End
   Begin VB.CommandButton Cancela 
      Caption         =   "Cancela"
      Height          =   315
      Left            =   5280
      TabIndex        =   10
      Top             =   2460
      Width           =   945
   End
   Begin VB.CommandButton Exclui 
      Caption         =   "Exclui"
      Height          =   315
      Left            =   5280
      TabIndex        =   9
      Top             =   1980
      Width           =   945
   End
   Begin VB.CommandButton Inclui 
      Caption         =   "Inclui"
      Height          =   315
      Left            =   5280
      TabIndex        =   7
      Top             =   1140
      Width           =   945
   End
   Begin VB.TextBox NOMPESPCP 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1920
      MaxLength       =   40
      TabIndex        =   1
      Top             =   180
      Width           =   4290
   End
   Begin VB.CommandButton BSel01 
      Caption         =   "?"
      Height          =   315
      Left            =   6300
      TabIndex        =   17
      Top             =   180
      Width           =   315
   End
   Begin VB.TextBox NOMCRG 
      Height          =   315
      Left            =   1140
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1020
      Width           =   2190
   End
   Begin VB.TextBox NOMDEP 
      Height          =   315
      Left            =   1140
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1440
      Width           =   3240
   End
   Begin VB.TextBox NOMPES 
      Height          =   315
      Left            =   1140
      MaxLength       =   40
      TabIndex        =   2
      Top             =   600
      Width           =   5070
   End
   Begin VB.CommandButton BSel02 
      Caption         =   "?"
      Height          =   315
      Left            =   6300
      TabIndex        =   18
      Top             =   600
      Width           =   315
   End
   Begin VB.TextBox NUMTLF 
      Height          =   315
      Left            =   1140
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1860
      Width           =   2190
   End
   Begin MSMask.MaskEdBox CODPES 
      Height          =   315
      Left            =   1140
      TabIndex        =   0
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   6
      Format          =   "000000"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox DIAANV 
      Height          =   315
      Left            =   1140
      TabIndex        =   6
      Top             =   2280
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   4
      Format          =   "00\/00"
      PromptChar      =   " "
   End
   Begin VB.Label Lb000 
      AutoSize        =   -1  'True
      Caption         =   "Pessoa"
      Height          =   195
      Left            =   315
      TabIndex        =   11
      Top             =   225
      Width           =   525
   End
   Begin VB.Label Lb001 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   195
      Left            =   315
      TabIndex        =   12
      Top             =   660
      Width           =   420
   End
   Begin VB.Label Lb002 
      AutoSize        =   -1  'True
      Caption         =   "Cargo"
      Height          =   195
      Left            =   315
      TabIndex        =   13
      Top             =   1080
      Width           =   420
   End
   Begin VB.Label Lb003 
      AutoSize        =   -1  'True
      Caption         =   "Depto"
      Height          =   195
      Left            =   315
      TabIndex        =   14
      Top             =   1500
      Width           =   435
   End
   Begin VB.Label Lb004 
      AutoSize        =   -1  'True
      Caption         =   "Telefone"
      Height          =   195
      Left            =   315
      TabIndex        =   15
      Top             =   1920
      Width           =   630
   End
   Begin VB.Label Lb005 
      AutoSize        =   -1  'True
      Caption         =   "Aniversário"
      Height          =   195
      Left            =   315
      TabIndex        =   16
      Top             =   2340
      Width           =   780
   End
End
Attribute VB_Name = "ATPPC01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TABPPC As csRS
Dim TABPES As csRS
 
Private Sub Altera_Click()
If Form_Ok() Then
   If TABPPC.Edit() Then
      FormToDb 1
      If TABPPC.Update() Then
         Form_Clear
      End If
   End If
End If
End Sub
 
 
Private Sub BSel01_Click()
If BRPES01(TABPES.RS) Then
   CODPES.Text = TABPES.RS!CODPES
   NOMPESPCP.Text = TABPES.RS!NOMPES
   CODPES.SetFocus
End If
End Sub
 
Private Sub BSel02_Click()
If oIsEmpty(CODPES.Text) Then
   ShowMessage "Informe a pessoa desejada ..."
 Else
   If BRPPC01(TABPPC.RS, (CODPES.Text)) Then
      DbToForm
      NOMPES.SetFocus
   End If
End If
End Sub
 
Private Sub Cancela_Click()
Unload Me
End Sub
 
 
Private Sub CODPES_GotFocus()
Inclui.Enabled = False
Altera.Enabled = False
Exclui.Enabled = False
End Sub
 
Private Sub CODPES_KeyPress(keyascii As Integer)
VldEditNum keyascii, "######", (CODPES.Text)
End Sub
 
Private Sub CODPES_LostFocus()
TABPES.RsSeek "=", CODPES.Text
If Not TABPES.Nomatch Then
   NOMPESPCP.Text = TABPES.RS!NOMPES
 Else
   NOMPESPCP.Text = ""
End If
End Sub
 
Private Sub Cst_CODPES(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(CODPES.Text) Then
      ShowMessage "Codigo da pessoa não informado..."
      Erro = True
      CODPES.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_DIAANV(Erro As Boolean)
If Not Erro Then
   If Not oIsEmpty(DIAANV.Text) Then
      If Not ddmmOk(CInt(DIAANV.Text)) Then
         ShowMessage "Dia do aniversário inválido..."
         DIAANV.SetFocus
         Erro = True
      End If
   End If
End If
End Sub
 
Private Sub Cst_NOMCRG(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(NOMCRG.Text) Then
      ShowMessage "Cargo não informado..."
      Erro = True
      NOMCRG.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_NOMDEP(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(NOMDEP.Text) Then
      ShowMessage "Nome do departamento não informado..."
      Erro = True
      NOMDEP.SetFocus
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
 
Private Sub Cst_NUMTLF(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(NUMTLF.Text) Then
      ShowMessage "Numero do telefone não informado..."
      Erro = True
      NUMTLF.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_Relacionamento(Erro As Boolean)
If Not Erro Then
   TABPES.RsSeek "=", CODPES.Text
   If TABPES.Nomatch Then
      Erro = True
      ShowMessage "Pessoa não cadastrado ..."
      CODPES.SetFocus
   End If
End If
End Sub
 
Private Sub DbToForm()
CODPES.Text = TABPPC.RS!CODPES
NOMPES.Text = TABPPC.RS!NOMPES
NOMCRG.Text = NullToSpc(TABPPC.RS!NOMCRG)
NOMDEP.Text = NullToSpc(TABPPC.RS!NOMDEP)
NUMTLF.Text = NullToSpc(TABPPC.RS!NUMTLF)
DIAANV.Text = NullToSpc(TABPPC.RS!DIAANV)
 
TABPES.RsSeek "=", TABPPC.RS!CODPES
If Not TABPES.Nomatch Then
   NOMPESPCP.Text = TABPES.RS!NOMPES
 Else
   NOMPESPCP.Text = ""
End If
End Sub
 
Private Sub DIAANV_KeyPress(keyascii As Integer)
VldEditNum keyascii, "####", (DIAANV.Text)
End Sub
 
Private Sub Exclui_Click()
If oConfirme("Confirma exclusao ?") Then
   If TABPPC.Delete() Then
      Form_Clear
   End If
End If
End Sub
 
Private Sub Form_Clear()
NOMPES.Text = ""
NOMCRG.Text = ""
NOMDEP.Text = ""
NUMTLF.Text = ""
DIAANV.Text = ""
NOMPES.SetFocus
Inclui.Enabled = False
Exclui.Enabled = False
Altera.Enabled = False
End Sub
 
Private Sub Form_KeyPress(keyascii As Integer)
If keyascii = vbKeyReturn Then
   keyascii = 0
   SendKeys Chr$(vbKeyTab)
End If
End Sub
 
Private Sub Form_Load()
HourGlassOn
Set TABPPC = oTABPPC()
Set TABPES = oTABPES()
Inclui.Enabled = False
Exclui.Enabled = False
Altera.Enabled = False
PositionForm Me, csFormRightBot
HourGlassOff
End Sub
 
Private Function Form_Ok() As Boolean
Dim Erro As Boolean
Erro = False
Cst_CODPES Erro
Cst_NOMPES Erro
Cst_NOMCRG Erro
Cst_NOMDEP Erro
Cst_NUMTLF Erro
Cst_DIAANV Erro
Cst_Relacionamento Erro
If Erro Then
   Form_Ok = False
 Else
   Form_Ok = True
End If
End Function
 
Private Sub Form_Unload(Cancel As Integer)
csRSClose TABPPC
csRSClose TABPES
End Sub
 
Private Sub FormToDb(OpIncAlt As Integer)
TABPPC.RS!CODPES = CODPES.Text
TABPPC.RS!NOMPES = NOMPES.Text
TABPPC.RS!NOMCRG = NOMCRG.Text
TABPPC.RS!NOMDEP = NOMDEP.Text
TABPPC.RS!NUMTLF = NUMTLF.Text
TABPPC.RS!DIAANV = SpcToNull(DIAANV.Text)
End Sub
 
Private Sub Inclui_Click()
If Form_Ok() Then
   If TABPPC.AddNew() Then
      FormToDb 0
      If TABPPC.Update() Then
         Form_Clear
      End If
   End If
End If
End Sub
 
Private Sub NOMPES_LostFocus()
If Not oIsEmpty(NOMPES.Text) Then
   TABPPC.RsSeek "=", CODPES.Text, NOMPES.Text
   If TABPPC.Nomatch Then
      Inclui.Enabled = True
      Altera.Enabled = False
      Exclui.Enabled = False
    Else
      DbToForm
      Inclui.Enabled = False
      Altera.Enabled = True
      Exclui.Enabled = True
   End If
End If
End Sub
 
 

