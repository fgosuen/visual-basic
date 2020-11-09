VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{ADD24EDC-ADC1-11D2-95D1-F7A835DD4948}#3.0#0"; "nslock15vb5.ocx"
Begin VB.MDIForm BMG001 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Gestão e Controle de Clientes"
   ClientHeight    =   3795
   ClientLeft      =   1455
   ClientTop       =   3945
   ClientWidth     =   12720
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin nslock15vb5.ActiveLock alock 
      Left            =   5850
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   820
      Password        =   "noitamiBMG"
      SoftwareName    =   "Gestão e Controle de Clientes"
      LiberationKeyLength=   16
      SoftwareCodeLength=   16
   End
   Begin Threed.SSPanel Panel1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12720
      _Version        =   65536
      _ExtentX        =   22437
      _ExtentY        =   979
      _StockProps     =   15
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSCommand lblregistro 
         Height          =   375
         Left            =   12915
         TabIndex        =   5
         Top             =   90
         Width           =   2040
         _Version        =   65536
         _ExtentX        =   3598
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Registrar Software"
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         Outline         =   0   'False
      End
      Begin Threed.SSCommand BT_PEDIDO 
         Height          =   450
         Left            =   1905
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   65
         Width           =   510
         _Version        =   65536
         _ExtentX        =   900
         _ExtentY        =   794
         _StockProps     =   78
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "BMG001.frx":0000
      End
      Begin Threed.SSCommand FINALIZA 
         Height          =   450
         Left            =   2655
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   65
         Width           =   510
         _Version        =   65536
         _ExtentX        =   900
         _ExtentY        =   794
         _StockProps     =   78
         Caption         =   "FIM"
         ForeColor       =   255
         RoundedCorners  =   0   'False
         Outline         =   0   'False
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   450
         Left            =   135
         TabIndex        =   4
         Top             =   65
         Width           =   510
         _Version        =   65536
         _ExtentX        =   900
         _ExtentY        =   794
         _StockProps     =   78
         RoundedCorners  =   0   'False
         Outline         =   0   'False
         Picture         =   "BMG001.frx":019A
      End
      Begin VB.Label lblaviso 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "PROGRAMA NÃO REGISTRADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   3825
         TabIndex        =   6
         Top             =   45
         Width           =   6015
      End
   End
   Begin Threed.SSPanel Message 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3420
      Width           =   12720
      _Version        =   65536
      _ExtentX        =   22437
      _ExtentY        =   661
      _StockProps     =   15
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   4
      BevelOuter      =   0
      BevelInner      =   1
      Alignment       =   1
      Begin VB.Timer Timer1 
         Interval        =   10000
         Left            =   6300
         Top             =   60
      End
   End
   Begin VB.Menu mnu_cadastros 
      Caption         =   "&Cadastros"
      Begin VB.Menu mnu_cadloja 
         Caption         =   "&Loja"
      End
      Begin VB.Menu mnu_cad_step05 
         Caption         =   "-"
      End
      Begin VB.Menu cad_estados 
         Caption         =   "&Estados"
      End
      Begin VB.Menu cad_cidades 
         Caption         =   "Ci&dades"
      End
      Begin VB.Menu t12 
         Caption         =   "-"
      End
      Begin VB.Menu cad_ramo_atv 
         Caption         =   "Ra&mos de Atividades"
      End
      Begin VB.Menu t1 
         Caption         =   "-"
      End
      Begin VB.Menu cad_clientes 
         Caption         =   "&Clientes"
      End
      Begin VB.Menu mnu_cad_limitecredito 
         Caption         =   "Limite de Crédito"
      End
   End
   Begin VB.Menu relatorios 
      Caption         =   "Relatóri&os"
      Begin VB.Menu rel_clientes 
         Caption         =   "&Clientes"
      End
      Begin VB.Menu t5 
         Caption         =   "-"
      End
      Begin VB.Menu rel_cidades 
         Caption         =   "Cidades"
      End
      Begin VB.Menu rel_Estados 
         Caption         =   "Estados"
      End
      Begin VB.Menu t14 
         Caption         =   "-"
      End
      Begin VB.Menu rel_ramos_atv 
         Caption         =   "Ra&mos de Atividades"
      End
   End
   Begin VB.Menu spool 
      Caption         =   "&Visualizador"
   End
   Begin VB.Menu mnu_gerencial 
      Caption         =   "&Gerencial"
      Begin VB.Menu mnu_alterasenha 
         Caption         =   "Alterar Senha"
      End
   End
   Begin VB.Menu fim 
      Caption         =   "&Fim"
   End
End
Attribute VB_Name = "BMG001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BTSel As Integer
 
Private Sub BR_FORNEC_Click()
ATPES01.ShowForm "F"
End Sub
 
Private Sub BT_CLIENTE_Click()
ATPES01.ShowForm "C"
End Sub
 
Private Sub BT_REPRES_Click()
ATPES01.ShowForm "R"
End Sub
 
Private Sub BT_SPOOL_Click()
Dim x
Screen.MousePointer = vbHourglass
x = Shell("viewer.exe cd .", 4)
Screen.MousePointer = vbDefault
End Sub
 
Private Sub BT_PEDIDO_Click()
Dim x
Screen.MousePointer = vbHourglass
x = Shell("viewer.exe bm .", 4)
Screen.MousePointer = vbDefault
End Sub

Private Sub cad_cidades_Click()
ATCID01.Show
End Sub
 
Private Sub cad_clientes_Click()
Load ATPES01
ATPES01.ShowForm "C"
End Sub
 
Private Sub cad_estados_Click()
ATUNF01.Show
End Sub
 
Private Sub cad_ramo_atv_Click()
ATRAT01.Show
End Sub
 
Private Sub Fim_Click()
Unload Me
End Sub
 
Private Sub FINALIZA_Click()
Unload Me
End Sub
 
Private Sub lblregistro_Click()
frmregistro.Show 1
End Sub

Private Sub MDIForm_Load()
If alock.LastRunDate > Now Then
   MsgBox "Ocorreu uma alteração na data do Sistema Operacional" _
   & vbCrLf & " O programa será encerrado.", vbOKOnly + vbCritical, "Erro de sistema"
   End
End If
If Not alock.RegisteredUser Then
   Me.Caption = Me.Caption & " (VERSÃO DE DEMONSTRAÇÃO)"
   BMG002.Caption = BMG002.Caption & " (VERSÃO DE DEMONSTRAÇÃO)"
Else
   Me.Caption = Me.Caption & " (VERSÃO REGISTRADA)"
   BMG002.Caption = BMG002.Caption & " (VERSÃO REGISTRADA)"
   lblaviso.Visible = False
   lblregistro.Enabled = False
End If
End Sub

Private Sub mnu_alterasenha_Click()
csPassWord.Show
End Sub

Private Sub mnu_cad_limitecredito_Click()
Dim f1 As New ATCRE01
f1.LimitarCredito = True
f1.Show
End Sub

Private Sub mnu_cadloja_Click()
ATLOJ01.Show
End Sub

Private Sub rel_cidades_Click()
Call RPCID01
End Sub
 
Private Sub rel_clientes_Click()
RPCLI01.Show
End Sub
 
Private Sub rel_Estados_Click()
Call RPUNF01
End Sub
 
Private Sub rel_ramos_atv_Click()
Call RPRAT01
End Sub
 
Private Sub spool_Click()
Dim x
Screen.MousePointer = vbHourglass
x = Shell("viewer.exe bm .", 4)
Screen.MousePointer = vbDefault
End Sub
 
Private Sub SSCommand1_Click()
Load ATPES01
ATPES01.ShowForm "C"
End Sub

Private Sub Timer1_Timer()
BMG001.Message.Caption = ""
Timer1.Enabled = False
BMG001.Message.Visible = False
End Sub
