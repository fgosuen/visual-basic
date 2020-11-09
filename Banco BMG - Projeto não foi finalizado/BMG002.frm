VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BMG002 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestão e Controle de Clientes"
   ClientHeight    =   3300
   ClientLeft      =   2895
   ClientTop       =   4290
   ClientWidth     =   8235
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3300
   ScaleWidth      =   8235
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2760
      Left            =   180
      Picture         =   "BMG002.frx":0000
      ScaleHeight     =   2700
      ScaleWidth      =   4680
      TabIndex        =   7
      Top             =   180
      Width           =   4740
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   3015
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13996
            Text            =   "Desenvolvido por Fernando Gosuen da Costa"
            TextSave        =   "Desenvolvido por Fernando Gosuen da Costa"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cancela 
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6705
      TabIndex        =   3
      Top             =   2340
      Width           =   1110
   End
   Begin VB.TextBox senha 
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
      IMEMode         =   3  'DISABLE
      Left            =   5385
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   2235
   End
   Begin VB.TextBox usuario 
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
      Left            =   5385
      MaxLength       =   10
      TabIndex        =   0
      Top             =   570
      Width           =   2235
   End
   Begin VB.CommandButton ok 
      Caption         =   "&Ok"
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
      Height          =   375
      Left            =   5445
      TabIndex        =   2
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
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
      Left            =   5430
      TabIndex        =   5
      Top             =   1095
      Width           =   465
   End
   Begin VB.Label label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário "
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
      Left            =   5445
      TabIndex        =   4
      Top             =   345
      Width           =   585
   End
End
Attribute VB_Name = "BMG002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SYSUSR As csRS
 
Private Sub Cancela_Click()
csRSClose SYSUSR
End
End Sub
 
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   SendKeys Chr$(vbKeyTab)
End If
End Sub
 
Private Sub Form_Load()
HourGlassOn

Set SYSUSR = oSYSUSR()

Picture1.Picture = LoadPicture(App.Path & "\bmg.bmp")

Me.Top = (Screen.Height - Me.Height) \ 2
Me.Left = (Screen.Width - Me.Width) \ 2

HourGlassOff
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
csRSClose SYSUSR
End Sub
 
Private Sub OK_Click()
If senha.Text = SYSUSR.RS!pswusr Then
   'BMG003.usuario.Caption = usuario.Text
   Unload Me
  Else
   ShowMessage "Senha inválida..."
End If
End Sub
 
Private Sub usuario_LostFocus()
If Not oIsEmpty(usuario.Text) Then
   usuario.Text = UCase$(usuario.Text)
   SYSUSR.RsSeek "=", usuario.Text
   If SYSUSR.Nomatch Then
      ok.Enabled = False
    Else
      ok.Enabled = True
    End If
End If
End Sub
