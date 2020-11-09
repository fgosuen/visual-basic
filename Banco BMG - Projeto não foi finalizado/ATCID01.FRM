VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ATCID01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Cidades"
   ClientHeight    =   2460
   ClientLeft      =   1635
   ClientTop       =   1650
   ClientWidth     =   5370
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
   LinkTopic       =   "ATCID01"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2460
   ScaleWidth      =   5370
   Begin VB.CommandButton PRXCID 
      Caption         =   "+"
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
      Left            =   2475
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   300
      Width           =   315
   End
   Begin VB.CommandButton BSel01 
      Caption         =   "?"
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
      Left            =   2070
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   300
      Width           =   315
   End
   Begin VB.TextBox NOMCID 
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
      Left            =   1350
      MaxLength       =   20
      TabIndex        =   3
      Top             =   690
      Width           =   3435
   End
   Begin VB.TextBox CODEST 
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
      Left            =   1350
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1125
      Width           =   495
   End
   Begin VB.TextBox NOMEST 
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
      Left            =   1935
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1125
      Width           =   2835
   End
   Begin VB.CommandButton Inclui 
      Caption         =   "Inclui"
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
      Left            =   495
      TabIndex        =   7
      Top             =   1905
      Width           =   945
   End
   Begin VB.CommandButton Altera 
      Caption         =   "Altera"
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
      Left            =   1635
      TabIndex        =   8
      Top             =   1905
      Width           =   945
   End
   Begin VB.CommandButton Exclui 
      Caption         =   "Exclui"
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
      Left            =   2760
      TabIndex        =   9
      Top             =   1905
      Width           =   945
   End
   Begin VB.CommandButton Cancela 
      Cancel          =   -1  'True
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
      Left            =   3900
      TabIndex        =   10
      Top             =   1905
      Width           =   945
   End
   Begin MSMask.MaskEdBox CODCID 
      Height          =   315
      Left            =   1350
      TabIndex        =   1
      Top             =   300
      Width           =   615
      _ExtentX        =   1085
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
   Begin VB.Label Lb000 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
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
      Left            =   705
      TabIndex        =   0
      Top             =   420
      Width           =   495
   End
   Begin VB.Label Lb001 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
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
      Left            =   705
      TabIndex        =   2
      Top             =   780
      Width           =   420
   End
   Begin VB.Label Lb002 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
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
      Left            =   705
      TabIndex        =   4
      Top             =   1185
      Width           =   495
   End
End
Attribute VB_Name = "ATCID01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TABPES As csRS
Dim TABLOJ As csRS
Dim TABCID As csRS
Dim TABUNF As csRS

Private Sub Altera_Click()
If FormOk() Then
   If TABCID.Edit() Then
      TABCID.RS!CODCID = CODCID.Text
      TABCID.RS!NOMCID = NOMCID.Text
      TABCID.RS!UNDFED = CODEST.Text
      If TABCID.Update() Then
         ClsForm
      End If
   End If
End If
End Sub

Private Sub BSel01_Click()
If BRCID01(TABCID.RS) Then
   CODCID.Text = TABCID.RS!CODCID
   NOMCID.Text = TABCID.RS!NOMCID
   CODEST.Text = TABCID.RS!UNDFED
   NOMEST.Text = NomeEstado(TABUNF, TABCID.RS!UNDFED)
   CODCID.SetFocus
End If
End Sub

Private Sub Cancela_Click()
Unload Me
End Sub

Private Sub ClsForm()
CODCID.Text = ""
NOMCID.Text = ""
CODEST.Text = ""
NOMEST.Text = ""
CODCID.SetFocus
Inclui.Enabled = False
Exclui.Enabled = False
Altera.Enabled = False
End Sub

Private Sub CODCID_KeyPress(keyascii As Integer)
VldEditNum keyascii, "####", (CODCID.Text)
End Sub

Private Sub CODCID_LostFocus()
If Not oIsEmpty(CODCID.Text) Then
   TABCID.RsSeek "=", CODCID.Text
   If TABCID.Nomatch Then
      Inclui.Enabled = True
      Altera.Enabled = False
      Exclui.Enabled = False
     Else
      CODCID.Text = TABCID.RS!CODCID
      NOMCID.Text = TABCID.RS!NOMCID
      CODEST.Text = TABCID.RS!UNDFED
      NOMEST.Text = NomeEstado(TABUNF, (CODEST.Text))
      Inclui.Enabled = False
      Altera.Enabled = True
      Exclui.Enabled = True
   End If
End If
End Sub

Private Sub CODEST_LostFocus()
CODEST.Text = UCase$(CODEST.Text)
ATCID01.NOMEST.Text = NomeEstado(TABUNF, (CODEST.Text))
End Sub

Private Sub Cst_CODCID(Erro As Integer)
If Erro Then
   Exit Sub
 Else
   If oIsEmpty(CODCID.Text) Then
      Erro = True
      ShowMessage "Código da Cidade não informado ..."
      CODCID.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_CODEST(Erro As Integer)
If Erro Then
   Exit Sub
 Else
   TABUNF.RsSeek "=", CODEST.Text
   If TABUNF.Nomatch Then
      Erro = True
      ShowMessage "Estado não cadastrado ..."
      ATCID01.CODEST.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_NOMCID(Erro As Integer)
If Erro Then
   Exit Sub
 Else
   If oIsEmpty(NOMCID.Text) Then
      Erro = True
      ShowMessage "Nome da Cidade não informado ..."
      ATCID01.NOMCID.SetFocus
   End If
End If
End Sub
 
Private Sub Exclui_Click()
If OkToDelete() Then
   If oConfirme("Confirma exclusão ?") Then
      If TABCID.Delete() Then
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
Set TABPES = oTABPES(): TABPES.Index = "IDXPES02"
Set TABLOJ = oTABLOJ(): TABLOJ.Index = "IDXLOJ02"
Set TABUNF = oTABUNF()
Set TABCID = oTABCID()
 
Inclui.Enabled = False
Exclui.Enabled = False
Altera.Enabled = False

Me.Top = (Screen.Height - Me.Height) \ 2
Me.Left = (Screen.Width - Me.Width) \ 2

HourGlassOff
End Sub

Private Sub Form_Unload(Cancel As Integer)
csRSClose TABPES
csRSClose TABLOJ
csRSClose TABCID
csRSClose TABUNF
End Sub

Private Function FormOk() As Integer
Dim Erro As Integer
Erro = False
Cst_CODCID Erro
Cst_NOMCID Erro
Cst_CODEST Erro
If Erro Then
   FormOk = False
 Else
   FormOk = True
End If
End Function

Private Sub Inclui_Click()
If FormOk() Then
   If TABCID.AddNew() Then
      TABCID.RS!CODCID = CODCID.Text
      TABCID.RS!NOMCID = NOMCID.Text
      TABCID.RS!UNDFED = CODEST.Text
      If TABCID.Update() Then
         ClsForm
      End If
   End If
End If
End Sub

Private Function OkToDelete() As Integer
Dim ok As Boolean
 
OkToDelete = True
ok = True
 
TABPES.RsSeek ">=", TABCID.RS!CODCID
If Not TABPES.Nomatch Then
   If TABCID.RS!CODCID = TABPES.RS!CODCID Then
      OkToDelete = False
      ShowMessage "Existem Clientes para esta cidade..."
      ok = False
   End If
End If
 
If ok Then
   TABLOJ.RsSeek "=", TABCID.RS!CODCID
   If Not TABLOJ.Nomatch Then
      OkToDelete = False
      ShowMessage "Existem Lojas para esta cidade..."
      ok = False
   End If
End If
End Function

Private Sub NOMCID_LostFocus()
NOMCID.Text = UCase(NOMCID.Text)
End Sub

Private Sub PRXCID_Click()
If TABCID.EOF And TABCID.BOF Then
   CODCID.Text = "1"
 Else
   TABCID.MoveLast
   CODCID.Text = Format$(TABCID.RS!CODCID + 1, "####")
End If
CODCID.SetFocus
End Sub
