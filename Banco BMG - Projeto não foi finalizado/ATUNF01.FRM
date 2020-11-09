VERSION 5.00
Begin VB.Form ATUNF01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Estados"
   ClientHeight    =   2055
   ClientLeft      =   4260
   ClientTop       =   3330
   ClientWidth     =   5130
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
   LinkTopic       =   "ATUNF01"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2055
   ScaleWidth      =   5130
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
      Left            =   2505
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   300
      Width           =   315
   End
   Begin VB.TextBox UNDFED 
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
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "*"
      Top             =   300
      Width           =   510
   End
   Begin VB.TextBox NOMUNF 
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
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   2
      Tag             =   "*"
      Top             =   690
      Width           =   2610
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
      Left            =   270
      TabIndex        =   3
      Top             =   1455
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
      Left            =   1500
      TabIndex        =   4
      Top             =   1470
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
      Left            =   2715
      TabIndex        =   5
      Top             =   1470
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
      Left            =   3915
      TabIndex        =   6
      Top             =   1455
      Width           =   945
   End
   Begin VB.Label Lb000 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sigla Estado"
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
      Left            =   660
      TabIndex        =   1
      Top             =   375
      Width           =   885
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
      Left            =   660
      TabIndex        =   8
      Top             =   720
      Width           =   420
   End
End
Attribute VB_Name = "ATUNF01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TABUNF As csRS
Dim TABCID As csRS
 
Private Sub Altera_Click()
If FormOk() Then
   If TABUNF.Edit() Then
      FormToDb
      If TABUNF.Update() Then
         ClsForm
      End If
   End If
End If
End Sub
 
Private Sub BSel01_Click()
If BRUNF01(TABUNF.RS) Then
   DbToForm
   UNDFED.SetFocus
End If
End Sub
 
Private Sub Cancela_Click()
Unload Me
End Sub
 
Private Sub ClsForm()
UNDFED.Text = ""
NOMUNF.Text = ""
UNDFED.SetFocus
Inclui.Enabled = False
Exclui.Enabled = False
Altera.Enabled = False
End Sub
 
Private Sub UNDFED_LostFocus()
If Not oIsEmpty(UNDFED.Text) Then
   UNDFED.Text = UCase$(UNDFED.Text)
   TABUNF.RsSeek "=", UNDFED.Text
   If TABUNF.Nomatch Then
      Inclui.Enabled = True
      Altera.Enabled = False
      Exclui.Enabled = False
     Else
      NOMUNF.Text = TABUNF.RS!NOMUNF
      Inclui.Enabled = False
      Altera.Enabled = True
      Exclui.Enabled = True
   End If
End If
End Sub
 
Private Sub Cst_UNDFED(Erro As Integer)
If Not Erro Then
  If oIsEmpty(ATUNF01.UNDFED.Text) Then
     ShowMessage "Informe codigo do estado ..."
     Erro = True
     ATUNF01.UNDFED.SetFocus
  End If
End If
End Sub
 
Private Sub CST_NOMUNF(Erro As Integer)
If Not Erro Then
  If oIsEmpty(ATUNF01.NOMUNF.Text) Then
     ShowMessage "Informe nome do estado ..."
     Erro = True
     ATUNF01.NOMUNF.SetFocus
  End If
End If
End Sub
 
Private Sub Exclui_Click()
If OkToDelete() Then
   If oConfirme("Confirma exclusao do estado ?") Then
      If TABUNF.Delete() Then
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
Set TABUNF = oTABUNF()
Set TABCID = oTABCID(): TABCID.Index = "IDXCID03"
 
Inclui.Enabled = False
Exclui.Enabled = False
Altera.Enabled = False

Me.Top = (Screen.Height - Me.Height) \ 2
Me.Left = (Screen.Width - Me.Width) \ 2

HourGlassOff
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
csRSClose TABUNF
csRSClose TABCID
End Sub
 
Private Function FormOk() As Integer
Dim Erro As Integer
Erro = False
Cst_UNDFED Erro
CST_NOMUNF Erro
If Erro Then
   FormOk = False
 Else
   FormOk = True
End If
End Function
 
Private Sub Inclui_Click()
If FormOk() Then
   If TABUNF.AddNew() Then
      FormToDb
      If TABUNF.Update() Then
         ClsForm
      End If
   End If
End If
End Sub
 
Private Function OkToDelete() As Integer
TABCID.RsSeek "=", TABUNF.RS!UNDFED
If TABCID.Nomatch Then
   OkToDelete = True
 Else
   OkToDelete = False
   ShowMessage "Estado possui cidades..."
End If
End Function
 
Private Sub DbToForm()
UNDFED.Text = TABUNF.RS!UNDFED
NOMUNF.Text = TABUNF.RS!NOMUNF
End Sub
 
Private Sub FormToDb()
TABUNF.RS!UNDFED = UNDFED.Text
TABUNF.RS!NOMUNF = NOMUNF.Text
End Sub
