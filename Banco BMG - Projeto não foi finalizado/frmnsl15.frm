VERSION 5.00
Begin VB.Form frmregistro 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Formulário de registro"
   ClientHeight    =   3915
   ClientLeft      =   4575
   ClientTop       =   5550
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   6180
   Begin VB.CommandButton cmdsair 
      Caption         =   "Sair do programa"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   3300
      Width           =   1575
   End
   Begin VB.CommandButton cmdregistrardepois 
      Caption         =   "Registrar Depois"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   3300
      Width           =   1695
   End
   Begin VB.CommandButton cmdregistraragora 
      Caption         =   "Registrar Agora"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3300
      Width           =   1575
   End
   Begin VB.TextBox txtcodigoliberacao 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2460
      Width           =   3375
   End
   Begin VB.TextBox txtcodigodoprograma 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1860
      Width           =   3375
   End
   Begin VB.TextBox txtdiasquefaltampararegistrar 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1260
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   3060
      Width           =   6255
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Liberação:"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   2565
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Código : "
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1935
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "dias para registrar o programa"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   1380
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Faltam"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1380
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   240
      Picture         =   "frmnsl15.frx":0000
      Top             =   240
      Width           =   555
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmnsl15.frx":03D5
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   840
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmregistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdregistraragora_Click()
If txtcodigoliberacao.Text = "" Then
   txtcodigoliberacao.SetFocus
   Exit Sub
End If

BMG001.alock.LiberationKey = txtcodigoliberacao.Text

If Not BMG001.alock.RegisteredUser Then
  MsgBox "Chave de LIBERAÇÃO INCORRETA", vbOKOnly + vbCritical, "Chave Liberação Incorreta"
   txtcodigoliberacao.SetFocus
Else
  MsgBox "REGISTRO EFETUADO COM SUCESSO !", vbExclamation, "Registro OK"
  BMG001.lblaviso.Visible = False
  BMG001.Caption = "Gestão e Controle de Clientes (VERSÃO REGISTRADA)"
  BMG001.lblregistro.Enabled = False
  Unload Me
End If
End Sub

Private Sub cmdregistrardepois_Click()
Unload Me
End Sub

Private Sub cmdsair_Click()
 End
End Sub

Private Sub Form_Load()
Dim diasQueFaltaParaRegistrar As Integer
diasQueFaltaParaRegistrar = 0
diasQueFaltaParaRegistrar = 15 - (BMG001.alock.UsedDays)
txtdiasquefaltampararegistrar.Text = diasQueFaltaParaRegistrar

If diasQueFaltaParaRegistrar <= 0 Then
   cmdregistrardepois.Enabled = False
End If

txtcodigodoprograma.Text = BMG001.alock.SoftwareCode
End Sub

Private Sub Label1_Click()

End Sub
