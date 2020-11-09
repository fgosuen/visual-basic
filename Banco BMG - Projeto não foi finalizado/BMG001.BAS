Attribute VB_Name = "BMG001M"
Option Explicit

Public G_RL_Empresa As String

Public Const PesJuridica = 0
Public Const PesFisica = 1

Public Sub Main()
Dim cDrive As String, cDir  As String
Dim D As Database, t  As RecordSet
Dim TABLOJ As csRS

cDrive = Mid$(App.Path, 1, 2)
cDir = Mid$(App.Path, 3)
 
ChDrive cDrive
ChDir cDir
 
AppSpoolPrefix = "bm"

Set TABLOJ = oTABLOJ()
TABLOJ.RsSeek "=", "1"
If Not TABLOJ.Nomatch Then
   G_RL_Empresa = TABLOJ.RS!NOMLOJ
Else
   G_RL_Empresa = "** Não cadastrado **"
End If

BMG001.Show
BMG002.Show vbModal
End Sub
 
Public Sub ShowMessage(cMsg As String)
BMG001.Message = cMsg: Beep
BMG001.Message.Visible = True
BMG001.Timer1.Enabled = True
End Sub

Public Function RetornaNOMCID(TABCID As csRS, nCodCid As Variant) As String
TABCID.RsSeek "=", nCodCid
If TABCID.Nomatch Then
   RetornaNOMCID = "Não Cadastrada"
Else
   RetornaNOMCID = TABCID.RS!NOMCID
End If
End Function
 
Public Function RetornaUNDFED(TABCID As csRS, nCodCid As Variant) As String
TABCID.RsSeek "=", nCodCid
If TABCID.Nomatch Then
   RetornaUNDFED = ""
Else
   RetornaUNDFED = TABCID.RS!UNDFED
End If
End Function
 
Public Function RetornaDESRAMATV(TABRAT As csRS, nCodRamAtv As Variant) As String
TABRAT.RsSeek "=", nCodRamAtv
If Not TABRAT.Nomatch Then
   RetornaDESRAMATV = NullToSpc(TABRAT.RS!DESRAMATV)
Else
   RetornaDESRAMATV = "Não Cadastrado"
End If
End Function

Public Function NomeEstado(TABUNF As csRS, nCodUNF As Variant) As String
If IsNull(nCodUNF) Then
   NomeEstado = ""
 Else
   TABUNF.RsSeek "=", nCodUNF
   If Not TABUNF.Nomatch Then
      NomeEstado = TABUNF.RS!NOMUNF
   Else
      NomeEstado = "** Não Cadastrado **"
   End If
End If
End Function

