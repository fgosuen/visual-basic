VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form ATRAT01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ramos de Atividade"
   ClientHeight    =   2340
   ClientLeft      =   3645
   ClientTop       =   4140
   ClientWidth     =   5235
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "ATRAT01"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2340
   ScaleWidth      =   5235
   Begin VB.CommandButton PRXCID 
      Caption         =   "+"
      Height          =   315
      Left            =   2070
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   315
   End
   Begin VB.CommandButton Altera 
      Caption         =   "Altera"
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   1636
      Width           =   945
   End
   Begin VB.CommandButton Cancela 
      Caption         =   "Cancela"
      Height          =   315
      Left            =   3735
      TabIndex        =   5
      Top             =   1636
      Width           =   945
   End
   Begin VB.CommandButton BSel01 
      Caption         =   "?"
      Height          =   315
      Left            =   1650
      TabIndex        =   6
      Top             =   480
      Width           =   315
   End
   Begin VB.TextBox DESRAMATV 
      Height          =   315
      Left            =   1140
      MaxLength       =   30
      TabIndex        =   1
      Top             =   855
      Width           =   3465
   End
   Begin VB.CommandButton Exclui 
      Caption         =   "Exclui"
      Height          =   315
      Left            =   2655
      TabIndex        =   4
      Top             =   1636
      Width           =   945
   End
   Begin VB.CommandButton Inclui 
      Caption         =   "Inclui"
      Height          =   315
      Left            =   465
      TabIndex        =   2
      Top             =   1636
      Width           =   945
   End
   Begin MSMask.MaskEdBox CODRAMATV 
      Height          =   315
      Left            =   1140
      TabIndex        =   0
      Top             =   480
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   3
      Format          =   "000"
      PromptChar      =   " "
   End
   Begin VB.Label Lb000 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   540
      Width           =   495
   End
   Begin VB.Label Lb001 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      Height          =   195
      Left            =   345
      TabIndex        =   8
      Top             =   915
      Width           =   720
   End
End
Attribute VB_Name = "ATRAT01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TABRAT As csRS

Private Sub Altera_Click()
If Form_Ok() Then
   If TABRAT.Edit() Then
      FormToDb 1
      If TABRAT.Update() Then
         Form_Clear
      End If
   End If
End If
End Sub

Private Sub BSel01_Click()
If BRRAT01(TABRAT.RS) Then
   DbToForm
   CODRAMATV.SetFocus
End If
End Sub

Private Sub Cancela_Click()
Unload Me
End Sub

Private Sub CODRAMATV_KeyPress(keyascii As Integer)
VldEditNum keyascii, "###", (CODRAMATV.Text)
End Sub

Private Sub CODRAMATV_LostFocus()
If Not oIsEmpty(CODRAMATV.Text) Then
   TABRAT.RsSeek "=", CODRAMATV.Text
   If TABRAT.Nomatch Then
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

Private Sub Cst_CODRAMATV(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(CODRAMATV.Text) Then
      ShowMessage "Código não informado..."
      Erro = True
      CODRAMATV.SetFocus
   End If
End If
End Sub

Private Sub Cst_DESRAMATV(Erro As Boolean)
If Not Erro Then
   If oIsEmpty(DESRAMATV.Text) Then
      ShowMessage "Descrição não informada..."
      Erro = True
      DESRAMATV.SetFocus
   End If
End If
End Sub

Private Sub DbToForm()
CODRAMATV.Text = TABRAT.RS!CODRAMATV
DESRAMATV.Text = TABRAT.RS!DESRAMATV
End Sub

Private Sub Exclui_Click()
If OkToDelete() Then
   If oConfirme("Confirma exclusão ?") Then
      If TABRAT.Delete() Then
         Form_Clear
      End If
   End If
End If
End Sub

Private Sub Form_Clear()
CODRAMATV.Text = ""
DESRAMATV.Text = ""
CODRAMATV.SetFocus
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

Set TABRAT = oTABRAT()

Inclui.Enabled = False
Exclui.Enabled = False
Altera.Enabled = False

Me.Top = (Screen.Height - Me.Height) \ 2
Me.Left = (Screen.Width - Me.Width) \ 2

HourGlassOff
End Sub

Private Function Form_Ok() As Boolean
Dim Erro As Boolean
Erro = False
Cst_CODRAMATV Erro
Cst_DESRAMATV Erro
If Erro Then
   Form_Ok = False
 Else
   Form_Ok = True
End If
End Function

Private Sub Form_Unload(Cancel As Integer)
csRSClose TABRAT
End Sub

Private Sub FormToDb(OpIncAlt As Integer)
TABRAT.RS!CODRAMATV = CODRAMATV.Text
TABRAT.RS!DESRAMATV = DESRAMATV.Text
End Sub

Private Sub Inclui_Click()
If Form_Ok() Then
   If TABRAT.AddNew() Then
      FormToDb 0
      If TABRAT.Update() Then
         Form_Clear
      End If
   End If
End If
End Sub

Private Function OkToDelete() As Boolean
Dim FlagOk As Boolean
Dim TABPES As csRS

Set TABPES = oTABPES(): TABPES.Index = "IDXPES06"

FlagOk = True

If FlagOk Then
   TABPES.RsSeek ">=", TABRAT.RS!CODRAMATV
   If Not TABPES.Nomatch Then
      If TABPES.RS!CODRAMATV = TABRAT.RS!CODRAMATV Then
         FlagOk = False
         ShowMessage "Existem Clientes com este Ramo de Atividade..."
      End If
   End If
End If

OkToDelete = FlagOk
csRSClose TABPES
End Function

Private Sub PRXCID_Click()
If TABRAT.EOF And TABRAT.BOF Then
   CODRAMATV.Text = "1"
 Else
   TABRAT.MoveLast
   CODRAMATV.Text = Format$(TABRAT.RS!CODRAMATV + 1, "###")
End If
CODRAMATV.SetFocus
End Sub
