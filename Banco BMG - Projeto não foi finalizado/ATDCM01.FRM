VERSION 5.00
Begin VB.Form ATDCM01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dados Complementares"
   ClientHeight    =   5430
   ClientLeft      =   2010
   ClientTop       =   3210
   ClientWidth     =   8985
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "ATDCM01"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5430
   ScaleWidth      =   8985
   Begin VB.TextBox REFPES 
      Height          =   1335
      Left            =   4740
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   3540
      Width           =   4155
   End
   Begin VB.TextBox REFCOM 
      Height          =   1300
      Left            =   4740
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   1920
      Width           =   4095
   End
   Begin VB.TextBox REFBAN 
      Height          =   1300
      Left            =   4740
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   300
      Width           =   4095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Emprego Anterior"
      Height          =   2235
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   4455
      Begin VB.TextBox TSVULTEMP 
         Height          =   285
         Left            =   180
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1740
         Width           =   1470
      End
      Begin VB.TextBox TLFULTEMP 
         Height          =   285
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1740
         Width           =   2265
      End
      Begin VB.TextBox LOCULTEMP 
         Height          =   285
         Left            =   180
         MaxLength       =   40
         TabIndex        =   6
         Top             =   540
         Width           =   4050
      End
      Begin VB.TextBox ENDULTEMP 
         Height          =   285
         Left            =   180
         MaxLength       =   60
         TabIndex        =   7
         Top             =   1140
         Width           =   4050
      End
      Begin VB.Label Lb011 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
         Height          =   195
         Left            =   1920
         TabIndex        =   23
         Top             =   1500
         Width           =   630
      End
      Begin VB.Label Lb010 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Lb009 
         AutoSize        =   -1  'True
         Caption         =   "Tempo"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label Lb008 
         AutoSize        =   -1  'True
         Caption         =   "Local"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   300
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Emprego Atual"
      Height          =   2235
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4455
      Begin VB.TextBox TSVEMP 
         Height          =   285
         Left            =   180
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1740
         Width           =   1530
      End
      Begin VB.TextBox TLFEMP 
         Height          =   285
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1740
         Width           =   2205
      End
      Begin VB.TextBox LOCEMP 
         Height          =   285
         Left            =   180
         MaxLength       =   40
         TabIndex        =   1
         Top             =   540
         Width           =   4110
      End
      Begin VB.TextBox ENDEMP 
         Height          =   285
         Left            =   180
         MaxLength       =   60
         TabIndex        =   2
         Top             =   1140
         Width           =   4095
      End
      Begin VB.Label Lb007 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
         Height          =   195
         Left            =   2040
         TabIndex        =   19
         Top             =   1500
         Width           =   630
      End
      Begin VB.Label Lb006 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   900
         Width           =   690
      End
      Begin VB.Label Lb005 
         AutoSize        =   -1  'True
         Caption         =   "Tempo"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label Lb004 
         AutoSize        =   -1  'True
         Caption         =   "Local "
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.CommandButton Altera 
      Caption         =   "Altera"
      Height          =   315
      Left            =   5040
      TabIndex        =   13
      Top             =   5040
      Width           =   945
   End
   Begin VB.CommandButton Cancela 
      Caption         =   "Cancela"
      Height          =   315
      Left            =   7500
      TabIndex        =   15
      Top             =   5040
      Width           =   945
   End
   Begin VB.CommandButton Exclui 
      Caption         =   "Exclui"
      Height          =   315
      Left            =   6360
      TabIndex        =   14
      Top             =   5040
      Width           =   945
   End
   Begin VB.Label Lb003 
      AutoSize        =   -1  'True
      Caption         =   "Ref. Pessoais"
      Height          =   195
      Left            =   4740
      TabIndex        =   26
      Top             =   3300
      Width           =   975
   End
   Begin VB.Label Lb002 
      AutoSize        =   -1  'True
      Caption         =   "Ref. Comerciais"
      Height          =   195
      Left            =   4740
      TabIndex        =   25
      Top             =   1680
      Width           =   1110
   End
   Begin VB.Label Lb001 
      AutoSize        =   -1  'True
      Caption         =   "Ref. Bancárias"
      Height          =   195
      Left            =   4740
      TabIndex        =   24
      Top             =   60
      Width           =   1050
   End
End
Attribute VB_Name = "ATDCM01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TABDCM As csRS
Dim CodPessoa As Long
Dim Incluir As Boolean
 
Public Sub ShowForm(Cod As Long)
Load Me
CodPessoa = Cod
TABDCM.RsSeek "=", CodPessoa
If Not TABDCM.Nomatch Then
   DbToForm
   Incluir = False
 Else
   Exclui.Enabled = False
   Incluir = True
End If
Me.Show vbModal
End Sub
 
Private Sub Altera_Click()
If Form_Ok() Then
   If Incluir Then
      If TABDCM.AddNew() Then
         FormToDb
         If TABDCM.Update() Then
            Unload Me
         End If
      End If
    Else
      If TABDCM.Edit() Then
         FormToDb
         If TABDCM.Update() Then
            Unload Me
         End If
      End If
   End If
End If
End Sub
 
Private Sub Cancela_Click()
Unload Me
End Sub
 
Private Sub DbToForm()
REFBAN.Text = NullToSpc(TABDCM.RS!REFBAN)
REFCOM.Text = NullToSpc(TABDCM.RS!REFCOM)
REFPES.Text = NullToSpc(TABDCM.RS!REFPES)
LOCEMP.Text = NullToSpc(TABDCM.RS!LOCEMP)
TSVEMP.Text = NullToSpc(TABDCM.RS!TSVEMP)
ENDEMP.Text = NullToSpc(TABDCM.RS!ENDEMP)
TLFEMP.Text = NullToSpc(TABDCM.RS!TLFEMP)
LOCULTEMP.Text = NullToSpc(TABDCM.RS!LOCULTEMP)
TSVULTEMP.Text = NullToSpc(TABDCM.RS!TSVULTEMP)
ENDULTEMP.Text = NullToSpc(TABDCM.RS!ENDULTEMP)
TLFULTEMP.Text = NullToSpc(TABDCM.RS!TLFULTEMP)
End Sub
 
Private Sub Exclui_Click()
If oConfirme("Confirma exclusao ?") Then
   If TABDCM.Delete() Then
      Unload Me
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
Set TABDCM = oTABDCM()
PositionForm Me, csFormRightBot
HourGlassOff
End Sub
 
Private Function Form_Ok() As Boolean
Dim Erro As Boolean
Erro = False
If Erro Then
   Form_Ok = False
 Else
   Form_Ok = True
End If
End Function
 
Private Sub Form_Unload(Cancel As Integer)
csRSClose TABDCM
End Sub
 
Private Sub FormToDb()
TABDCM.RS!CODPES = CodPessoa
TABDCM.RS!REFBAN = SpcToNull(REFBAN.Text)
TABDCM.RS!REFCOM = SpcToNull(REFCOM.Text)
TABDCM.RS!REFPES = SpcToNull(REFPES.Text)
TABDCM.RS!LOCEMP = SpcToNull(LOCEMP.Text)
TABDCM.RS!TSVEMP = SpcToNull(TSVEMP.Text)
TABDCM.RS!ENDEMP = SpcToNull(ENDEMP.Text)
TABDCM.RS!TLFEMP = SpcToNull(TLFEMP.Text)
TABDCM.RS!LOCULTEMP = SpcToNull(LOCULTEMP.Text)
TABDCM.RS!TSVULTEMP = SpcToNull(TSVULTEMP.Text)
TABDCM.RS!ENDULTEMP = SpcToNull(ENDULTEMP.Text)
TABDCM.RS!TLFULTEMP = SpcToNull(TLFULTEMP.Text)
End Sub
