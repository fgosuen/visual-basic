VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ATCRE01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Atualiza Limite de Crédito"
   ClientHeight    =   1995
   ClientLeft      =   3645
   ClientTop       =   3300
   ClientWidth     =   6195
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "ATCRE01"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1995
   ScaleWidth      =   6195
   Begin VB.CommandButton Altera 
      Caption         =   "Altera"
      Height          =   315
      Left            =   2715
      TabIndex        =   3
      Top             =   1305
      Width           =   975
   End
   Begin VB.CommandButton Cancela 
      Caption         =   "Cancela"
      Height          =   315
      Left            =   4275
      TabIndex        =   4
      Top             =   1305
      Width           =   975
   End
   Begin VB.TextBox NOMCLI 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2055
      MaxLength       =   40
      TabIndex        =   1
      Top             =   225
      Width           =   3555
   End
   Begin VB.CommandButton BSel01 
      Caption         =   "?"
      Height          =   315
      Left            =   5655
      TabIndex        =   5
      Top             =   225
      Width           =   315
   End
   Begin MSMask.MaskEdBox CODCLI 
      Height          =   315
      Left            =   1275
      TabIndex        =   0
      Top             =   225
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   6
      Format          =   "000000"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox LIMCRE 
      Height          =   315
      Left            =   1275
      TabIndex        =   2
      Top             =   645
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   6
      Format          =   "##0.00"
      PromptChar      =   " "
   End
   Begin VB.Label Lb000 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      Height          =   195
      Left            =   255
      TabIndex        =   7
      Top             =   285
      Width           =   480
   End
   Begin VB.Label Lb003 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Limite Crédito"
      Height          =   195
      Left            =   255
      TabIndex        =   6
      Top             =   705
      Width           =   945
   End
End
Attribute VB_Name = "ATCRE01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TABCLI As csRS
Dim TABPRM As csRS

Public LimitarCredito As Boolean
 
Private Sub Altera_Click()
If Form_Ok() Then
   If TABCLI.Edit() Then
      FormToDb 1
      If TABCLI.Update() Then
         Form_Clear
      End If
   End If
End If
End Sub
 
Private Sub BSel01_Click()
If BRCLI01(TABCLI.RS) Then
   DbToForm
   CODCLI.SetFocus
End If
End Sub
 
Private Sub Cancela_Click()
Unload Me
End Sub
 
Private Sub CODCLI_KeyPress(KeyAscii As Integer)
VldEditNum KeyAscii, "######", (CODCLI.Text)
End Sub
 
Private Sub CODCLI_LostFocus()
TABCLI.RsSeek "=", CODCLI.Text
If Not TABCLI.Nomatch Then
   Altera.Enabled = True
   DbToForm
 Else
   Altera.Enabled = False
End If
End Sub
 
Private Sub Cst_CODCLI(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(CODCLI.Text) Then
      ShowMessage "Informe o codido do cliente ..."
      Erro = True
      CODCLI.SetFocus
   End If
End If
End Sub
 
Private Sub Cst_LIMCRE(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(LIMCRE.Text) Then
      ShowMessage "Informe valor do limite de crédito..."
      Erro = True
      LIMCRE.SetFocus
    Else
      If LimitarCredito Then
         If CCur(LIMCRE.Text) > NullToZero(TABPRM.RS!VLRMAXCRE) Then
            ShowMessage "Valor do limite de crédito inválido..."
            Erro = True
            LIMCRE.SetFocus
         End If
      End If
   End If
End If
End Sub
 
Private Sub DbToForm()
CODCLI.Text = TABCLI.RS!CODCLI
NOMCLI.Text = TABCLI.RS!NOMCLI
LIMCRE.Text = NullToZero(TABCLI.RS!LIMCRE)
End Sub
 
Private Sub Form_Clear()
CODCLI.Text = ""
NOMCLI.Text = ""
LIMCRE.Text = ""
CODCLI.SetFocus
Altera.Enabled = False
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
Set TABPRM = oTABPRM()

Altera.Enabled = False

Me.Top = (Screen.Height - Me.Height) \ 2
Me.Left = (Screen.Width - Me.Width) \ 2

HourGlassOff
End Sub
 
Private Function Form_Ok() As Boolean
Dim Erro As Boolean
Erro = False
Cst_CODCLI Erro
Cst_LIMCRE Erro
If Erro Then
   Form_Ok = False
 Else
   Form_Ok = True
End If
End Function
 
Private Sub Form_Unload(Cancel As Integer)
csRSClose TABCLI
csRSClose TABPRM
End Sub
 
Private Sub FormToDb(OpIncAlt As Integer)
TABCLI.RS!LIMCRE = LIMCRE.Text
End Sub
 
Private Sub LIMCRE_KeyPress(KeyAscii As Integer)
VldEditNum KeyAscii, "######.##", (LIMCRE.Text)
End Sub
 

