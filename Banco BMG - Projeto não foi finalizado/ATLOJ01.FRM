VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ATLOJ01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro da Loja"
   ClientHeight    =   4035
   ClientLeft      =   2880
   ClientTop       =   4980
   ClientWidth     =   6255
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "ATLOJ01"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4035
   ScaleWidth      =   6255
   Begin VB.TextBox NOMLOJ 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   1
      Top             =   540
      Width           =   4455
   End
   Begin VB.TextBox NOMCID 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   3
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox DESEND 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1380
      Width           =   4485
   End
   Begin VB.TextBox NOMBAI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1800
      Width           =   2505
   End
   Begin VB.TextBox NUMTLF 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2220
      Width           =   1815
   End
   Begin VB.TextBox NUMFAX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      MaxLength       =   20
      TabIndex        =   8
      Top             =   2220
      Width           =   1815
   End
   Begin VB.TextBox INSEST 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   10
      Top             =   3060
      Width           =   2490
   End
   Begin VB.CommandButton Altera 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3510
      TabIndex        =   11
      Top             =   3570
      Width           =   945
   End
   Begin VB.CommandButton Cancela 
      Caption         =   "Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5085
      TabIndex        =   12
      Top             =   3570
      Width           =   945
   End
   Begin MSMask.MaskEdBox CODLOJ 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "###"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox CODCID 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NUMCEP 
      Height          =   315
      Left            =   4980
      TabIndex        =   6
      Top             =   1800
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00000\-000"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NUMCGC 
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Top             =   2640
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00\-000\-000\/0000\-00"
      PromptChar      =   " "
   End
   Begin VB.Label Lb000 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   13
      Top             =   225
      Width           =   495
   End
   Begin VB.Label Lb001 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome da Loja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   14
      Top             =   600
      Width           =   990
   End
   Begin VB.Label Lb002 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   15
      Top             =   1020
      Width           =   495
   End
   Begin VB.Label Lb004 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   16
      Top             =   1440
      Width           =   690
   End
   Begin VB.Label Lb005 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   17
      Top             =   1860
      Width           =   405
   End
   Begin VB.Label Lb006 
      AutoSize        =   -1  'True
      Caption         =   "CEP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4620
      TabIndex        =   18
      Top             =   1860
      Width           =   315
   End
   Begin VB.Label Lb007 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   19
      Top             =   2340
      Width           =   630
   End
   Begin VB.Label Lb008 
      AutoSize        =   -1  'True
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   20
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Lb009 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.G.C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   21
      Top             =   2700
      Width           =   420
   End
   Begin VB.Label Lb010 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inscricao Estadual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   22
      Top             =   3120
      Width           =   1305
   End
End
Attribute VB_Name = "ATLOJ01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TABLOJ As csRS
Dim TABCID As csRS

Private Sub Altera_Click()
If FormOk() Then
   If TABLOJ.Edit() Then
      FormToDb
      If TABLOJ.Update() Then
         ClsForm
      End If
   End If
End If
End Sub

Private Sub BSel01_Click()
If BRLOJ01(TABLOJ.RS) Then
   DbToForm
   CODLOJ.SetFocus
End If
End Sub

Private Sub Cancela_Click()
Unload Me
End Sub

Private Sub ClsForm()
'CODLOJ.Text = ""
'NOMLOJ.Text = ""
'CODCID.Text = ""
'NOMCID.Text = ""
'DESEND.Text = ""
'NOMBAI.Text = ""
'NUMCEP.Text = ""
'NUMTLF.Text = ""
'NUMFAX.Text = ""
'NUMCGC.Text = ""
'INSEST.Text = ""
CODLOJ.SetFocus
'Inclui.Enabled = False
'Exclui.Enabled = False
'Altera.Enabled = False
iMsgBox "Registro atualizado !"
End Sub

Private Sub CODCID_KeyPress(keyascii As Integer)
VldEditNum keyascii, "####", (CODCID.Text)
End Sub

Private Sub CODCID_LostFocus()
TABCID.RsSeek "=", CODCID.Text
If TABCID.Nomatch Then
   NOMCID.Text = ""
 Else
   NOMCID.Text = TABCID.RS!NOMCID
End If
End Sub

Private Sub CODLOJ_KeyPress(keyascii As Integer)
VldEditNum keyascii, "###", (CODLOJ.Text)
End Sub

Private Sub CODLOJ_LostFocus()
If Not oIsEmpty(CODLOJ.Text) Then
   TABLOJ.RsSeek "=", "1"  '(CODLOJ.Text)
   If TABLOJ.Nomatch Then
      'Inclui.Enabled = True
      Altera.Enabled = False
      'Exclui.Enabled = False
     Else
      DbToForm
      'Inclui.Enabled = False
      Altera.Enabled = True
      'Exclui.Enabled = True
   End If
End If
End Sub

Private Sub Cst_CODCID(Erro As Integer)
If Not Erro Then
   TABCID.RsSeek "=", Val(CODCID.Text)
   If TABCID.Nomatch Then
      Erro = True
      ShowMessage "Cidade não cadastrada ..."
      CODCID.SetFocus
   End If
End If
End Sub

Private Sub Cst_CODLOJ(Erro As Integer)
If Not Erro Then
   If oIsEmpty(CODLOJ.Text) Then
      Erro = True
      ShowMessage "Informação inválida ..."
      CODLOJ.SetFocus
   End If
End If
End Sub

Private Sub Cst_DESEND(Erro As Integer)
If Not Erro Then
   If oIsEmpty(DESEND.Text) Then
      Erro = True
      ShowMessage "Informacao invalida ..."
      DESEND.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_INSEST(Erro As Integer)
If Not Erro Then
   If oIsEmpty(INSEST.Text) Then
      Erro = True
      ShowMessage "Informacao invalida ..."
      INSEST.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_NOMBAI(Erro As Integer)
If Not Erro Then
   If oIsEmpty(NOMBAI.Text) Then
      Erro = True
      ShowMessage "Informacao invalida ..."
      NOMBAI.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_NOMLOJ(Erro As Integer)
If Not Erro Then
   If oIsEmpty(NOMLOJ.Text) Then
      Erro = True
      ShowMessage "Informacao invalida ..."
      NOMLOJ.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_NUMCEP(Erro As Integer)
If Not Erro Then
   If oIsEmpty(NUMCEP.Text) Then
      Erro = True
      ShowMessage "Informacao invalida ..."
      NUMCEP.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_NUMCGC(Erro As Integer)
If Not Erro Then
   If oIsEmpty(NUMCGC.Text) Then
      Erro = True
      ShowMessage "Informacao invalida ..."
      NUMCGC.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_NUMFAX(Erro As Integer)
If Not Erro Then
   If oIsEmpty(NUMFAX.Text) Then
      Erro = True
      ShowMessage "Informacao invalida ..."
      NUMFAX.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_NUMTLF(Erro As Integer)
If Not Erro Then
   If oIsEmpty(NUMTLF.Text) Then
      Erro = True
      ShowMessage "Informacao invalida ..."
      NUMTLF.SetFocus
   End If
End If
End Sub
 
Private Sub Exclui_Click()
If OkToDelete() Then
   If oConfirme("Confirma exclusao ?") Then
      If TABLOJ.Delete() Then
         ClsForm
      End If
   End If
End If
End Sub
 
Private Sub Form_KeyPress(keyascii As Integer)
If keyascii = vbKeyReturn Then
   keyascii = 0
   SendKeys Chr$(vbKeyTab)
End If
End Sub
 
Private Sub Form_Load()
HourGlassOn
Set TABLOJ = oTABLOJ()
Set TABCID = oTABCID()

TABLOJ.RsSeek "=", "1"
If Not TABLOJ.Nomatch Then
   DbToForm
End If

Me.Top = (Screen.Height - Me.Height) \ 2
Me.Left = (Screen.Width - Me.Width) \ 2

HourGlassOff
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
csRSClose TABLOJ
csRSClose TABCID
End Sub
 
Private Function FormOk() As Integer
Dim Erro As Integer
Erro = False
Cst_CODLOJ Erro
Cst_NOMLOJ Erro
Cst_CODCID Erro
Cst_DESEND Erro
Cst_NOMBAI Erro
Cst_NUMCEP Erro
Cst_NUMTLF Erro
Cst_NUMFAX Erro
Cst_NUMCGC Erro
Cst_INSEST Erro
If Erro Then
   FormOk = False
 Else
   FormOk = True
End If
End Function
 
Private Sub Inclui_Click()
If FormOk() Then
   If TABLOJ.AddNew() Then
      FormToDb
      If TABLOJ.Update() Then
         ClsForm
      End If
   End If
End If
End Sub
 
Private Sub NUMCEP_KeyPress(keyascii As Integer)
VldEditNum keyascii, "########", (NUMCEP.Text)
End Sub
 
Private Sub NUMCGC_KeyPress(keyascii As Integer)
VldEditNum keyascii, "##############", (NUMCGC.Text)
End Sub
 
Private Function OkToDelete() As Integer
OkToDelete = True
End Function
 
Private Sub DbToForm()
CODLOJ.Text = TABLOJ.RS!CODLOJ
NOMLOJ.Text = TABLOJ.RS!NOMLOJ
CODCID.Text = TABLOJ.RS!CODCID
DESEND.Text = TABLOJ.RS!DESEND
NOMBAI.Text = TABLOJ.RS!NOMBAI
NUMCEP.Text = TABLOJ.RS!NUMCEP
NUMTLF.Text = TABLOJ.RS!NUMTLF
NUMFAX.Text = TABLOJ.RS!NUMFAX
NUMCGC.Text = TABLOJ.RS!NUMCGC
INSEST.Text = TABLOJ.RS!INSEST
 
TABCID.RsSeek "=", TABLOJ.RS!CODCID
If TABCID.Nomatch Then
   NOMCID.Text = ""
 Else
   NOMCID.Text = TABCID.RS!NOMCID
End If
End Sub
 
Private Sub FormToDb()
TABLOJ.RS!CODLOJ = CODLOJ.Text
TABLOJ.RS!NOMLOJ = NOMLOJ.Text

G_RL_Empresa = TABLOJ.RS!NOMLOJ

TABLOJ.RS!CODCID = CODCID.Text
TABLOJ.RS!DESEND = DESEND.Text
TABLOJ.RS!NOMBAI = NOMBAI.Text
TABLOJ.RS!NUMCEP = NUMCEP.Text
TABLOJ.RS!NUMTLF = NUMTLF.Text
TABLOJ.RS!NUMFAX = NUMFAX.Text
TABLOJ.RS!NUMCGC = NUMCGC.Text
TABLOJ.RS!INSEST = INSEST.Text
End Sub
 

