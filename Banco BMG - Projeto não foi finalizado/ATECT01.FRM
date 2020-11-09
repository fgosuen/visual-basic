VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ATECT01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Endereço de Coleta"
   ClientHeight    =   4095
   ClientLeft      =   1605
   ClientTop       =   2415
   ClientWidth     =   6360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4095
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton INCLUI 
      Caption         =   "&Inclui"
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   3540
      Width           =   1035
   End
   Begin VB.CommandButton ALTERA 
      Caption         =   "&Altera"
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   3540
      Width           =   1035
   End
   Begin VB.CommandButton SELECIONA 
      Caption         =   "&Seleciona"
      Height          =   375
      Left            =   300
      TabIndex        =   11
      Top             =   3540
      Width           =   1035
   End
   Begin VB.CommandButton SelCODCID 
      Caption         =   "?"
      Height          =   285
      Left            =   5760
      TabIndex        =   25
      Top             =   1200
      Width           =   285
   End
   Begin VB.TextBox UNDFED 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5160
      TabIndex        =   7
      Top             =   1200
      Width           =   555
   End
   Begin VB.CommandButton Cancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4980
      TabIndex        =   15
      Top             =   3540
      Width           =   1035
   End
   Begin VB.CommandButton EXCLUI 
      Caption         =   "&Exclui"
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   3540
      Width           =   1035
   End
   Begin VB.TextBox CMPOBS 
      Height          =   915
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2340
      Width           =   5775
   End
   Begin VB.TextBox NUMFAX 
      Height          =   285
      Left            =   4260
      MaxLength       =   15
      TabIndex        =   9
      Top             =   1560
      Width           =   1755
   End
   Begin VB.TextBox NUMTLF 
      Height          =   285
      Left            =   1260
      MaxLength       =   15
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox NOMCID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1860
      TabIndex        =   6
      Top             =   1200
      Width           =   3195
   End
   Begin VB.TextBox NOMBAI 
      Height          =   285
      Left            =   1260
      MaxLength       =   30
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox DESEND 
      Height          =   285
      Left            =   1260
      MaxLength       =   40
      TabIndex        =   2
      Top             =   480
      Width           =   4815
   End
   Begin VB.TextBox NOMTRP 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2100
      MaxLength       =   40
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin MSMask.MaskEdBox CODCID 
      Height          =   285
      Left            =   1260
      TabIndex        =   5
      Top             =   1200
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   503
      _Version        =   393216
      Format          =   "0000"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox NUMCEP 
      Height          =   285
      Left            =   5160
      TabIndex        =   4
      Top             =   840
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      _Version        =   393216
      Format          =   "00000\-000"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox CODTRP 
      Height          =   285
      Left            =   1260
      TabIndex        =   0
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   503
      _Version        =   393216
      ForeColor       =   12632256
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Ctrl + Enter para próxima linha"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3900
      TabIndex        =   24
      Top             =   2100
      Width           =   2115
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Observações"
      Height          =   195
      Left            =   300
      TabIndex        =   23
      Top             =   2100
      Width           =   945
   End
   Begin VB.Label Label7 
      Caption         =   "Fax"
      Height          =   195
      Left            =   3960
      TabIndex        =   22
      Top             =   1620
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Telefone"
      Height          =   195
      Left            =   540
      TabIndex        =   21
      Top             =   1620
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cidade"
      Height          =   195
      Left            =   660
      TabIndex        =   20
      Top             =   1260
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "CEP"
      Height          =   195
      Left            =   4740
      TabIndex        =   19
      Top             =   900
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Bairro"
      Height          =   195
      Left            =   780
      TabIndex        =   18
      Top             =   900
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Endereço"
      Height          =   195
      Left            =   480
      TabIndex        =   17
      Top             =   540
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Transportadora"
      Enabled         =   0   'False
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   180
      Width           =   1080
   End
End
Attribute VB_Name = "ATECT01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TABCID As csRS
Dim TABECT As csRS
 
Public Sub Form_Clear()
CODCID.Text = ""
DESEND.Text = ""
NOMBAI.Text = ""
NUMCEP.Text = ""
NUMTLF.Text = ""
NUMFAX.Text = ""
CMPOBS.Text = ""
UNDFED.Text = ""
NOMCID.Text = ""
Altera.Enabled = False
Exclui.Enabled = False
Inclui.Enabled = True
DESEND.SetFocus
End Sub
 
Private Sub cancelar_Click()
Unload Me
End Sub
 
Private Sub CODCID_KeyPress(keyascii As Integer)
VldEditNum keyascii, "####", (CODCID.Text)
End Sub
 
Private Sub CODCID_LostFocus()
If Not oIsEmpty(CODCID.Text) Then
   TABCID.RsSeek "=", Val(CODCID.Text)
   If Not TABCID.Nomatch Then
      NOMCID.Text = NullToSpc(TABCID.RS!NOMCID)
      UNDFED.Text = NullToSpc(TABCID.RS!UNDFED)
    Else
      NOMCID.Text = "Não Cadastrada"
      UNDFED.Text = ""
   End If
Else
   NOMCID.Text = ""
   UNDFED.Text = ""
End If
End Sub
 
 
Private Sub Exclui_Click()
If oConfirme("Confirma exclusão ?") Then
   If TABECT.Delete() Then
      Form_Clear
   End If
End If
End Sub
 
Private Sub Form_Load()
HourGlassOn
Set TABCID = oTABCID()
Set TABECT = oTABECT()
Altera.Enabled = False
Exclui.Enabled = False
PositionForm Me, csFormRightBot
HourGlassOff
End Sub
 
Private Sub Form_KeyPress(keyascii As Integer)
If keyascii = vbKeyReturn Then
   keyascii = 0
   SendKeys Chr$(vbKeyTab)
End If
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
csRSClose TABCID
csRSClose TABECT
End Sub
 
Private Sub Inclui_Click()
Dim oTrans As New csTransaction
If FormOk() Then
   HourGlassOn
   oTrans.BegTrans
   If oTrans.AddNew(TABECT.RS) Then
      FormToDb
      oTrans.Update TABECT.RS
   End If
   oTrans.EndTrans
   HourGlassOff
 
   If oTrans.ErrorCode = 0 Then
      Form_Clear
   End If
End If
 
End Sub
 
Private Sub NUMCEP_KeyPress(keyascii As Integer)
VldEditNum keyascii, "########", (NUMCEP.Text)
End Sub
 
Private Sub Altera_Click()
Dim oTrans As New csTransaction
If FormOk() Then
   HourGlassOn
   oTrans.BegTrans
   If oTrans.Edit(TABECT.RS) Then
      FormToDb
      oTrans.Update TABECT.RS
   End If
   oTrans.EndTrans
   HourGlassOff
   If oTrans.ErrorCode = 0 Then
      Form_Clear
   End If
End If
End Sub
 
Private Sub SelCODCID_Click()
If BRCID01(TABCID.RS) Then
   CODCID.Text = NullToSpc(TABCID.RS!CODCID)
   CODCID_LostFocus
End If
CODCID.SetFocus
End Sub
 
Public Sub DbToForm()
CODTRP.Text = TABECT.RS!CODTRP
CODCID.Text = NullToSpc(TABECT.RS!CODCID)
DESEND.Text = NullToSpc(TABECT.RS!DESEND)
NOMBAI.Text = NullToSpc(TABECT.RS!NOMBAI)
NUMCEP.Text = NullToSpc(TABECT.RS!NUMCEP)
NUMTLF.Text = NullToSpc(TABECT.RS!NUMTLF)
NUMFAX.Text = NullToSpc(TABECT.RS!NUMFAX)
CMPOBS.Text = NullToSpc(TABECT.RS!CMPOBS)
CODCID_LostFocus
End Sub
 
Private Sub FormToDb()
TABECT.RS!CODTRP = SpcToNull(CODTRP.Text)
TABECT.RS!CODCID = SpcToNull(CODCID.Text)
TABECT.RS!DESEND = SpcToNull(DESEND.Text)
TABECT.RS!NOMBAI = SpcToNull(NOMBAI.Text)
TABECT.RS!NUMCEP = SpcToNull(NUMCEP.Text)
TABECT.RS!NUMTLF = SpcToNull(NUMTLF.Text)
TABECT.RS!NUMFAX = SpcToNull(NUMFAX.Text)
TABECT.RS!CMPOBS = SpcToNull(CMPOBS.Text)
End Sub
Private Function FormOk() As Boolean
Dim Erro As Boolean
Erro = False
Cst_CODCID Erro
Cst_DESEND Erro
Cst_NOMBAI Erro
Cst_NUMCEP Erro
Cst_NUMTLF Erro
Cst_NUMFAX Erro
If Not Erro Then
   FormOk = True
Else
   FormOk = False
End If
End Function
 
Private Sub Cst_CODCID(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(CODCID.Text) Then
      ShowMessage "Informe o codigo da ciade..."
      Erro = True
      CODCID.SetFocus
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
      ShowMessage "Endereco nao informado ..."
      Erro = True
      DESEND.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_NOMBAI(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(NOMBAI.Text) Then
      ShowMessage "Informe nome do bairro..."
      Erro = True
      NOMBAI.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_NUMTLF(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(NUMTLF.Text) Then
      ShowMessage "Informe numero do telefone ..."
      Erro = True
      NUMTLF.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_NUMFAX(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(NUMFAX.Text) Then
      ShowMessage "Informe numero do FAX ..."
      Erro = True
      NUMFAX.SetFocus
   End If
End If
End Sub
 
Private Sub SELECIONA_Click()
If BRECT01(TABECT.RS, (CODTRP.Text)) Then
   DbToForm
   Inclui.Enabled = False
   Altera.Enabled = True
   Exclui.Enabled = True
 Else
   Inclui.Enabled = True
   Altera.Enabled = False
   Exclui.Enabled = False
End If
DESEND.SetFocus
End Sub
