Attribute VB_Name = "mDBPES"
Option Explicit
 
Private Const cteDBPES = "DBPES.MDB"
 
Public Function oDBPES() As csDB
Set oDBPES = csDbOpen(cteDBPES)
End Function

Public Function oTABPRM(Optional cSql As Variant) As csRS
Set oTABPRM = Abre(cSql, "TABPRM")
End Function

Public Function oTABECT(Optional cSql As Variant) As csRS
Set oTABECT = Abre(cSql, "TABECT", "IDXECT01")
End Function
 
Public Function oTABCID(Optional cSql As Variant) As csRS
Set oTABCID = Abre(cSql, "TABCID", "IDXCID01")
End Function
 
Public Function oTABLOJ(Optional cSql As Variant) As csRS
Set oTABLOJ = Abre(cSql, "TABLOJ", "IDXLOJ01")
End Function
 
Public Function oTABCLI(Optional cSql As Variant) As csRS
Set oTABCLI = Abre(cSql, "TABCLI", "IDXCLI01")
End Function
 
Public Function oTABDCM(Optional cSql As Variant) As csRS
Set oTABDCM = Abre(cSql, "TABDCM", "IDXDCM01")
End Function
 
Public Function oTABFOR(Optional cSql As Variant) As csRS
Set oTABFOR = Abre(cSql, "TABFOR", "IDXFOR01")
End Function
 
Public Function oTABPES(Optional cSql As Variant) As csRS
Set oTABPES = Abre(cSql, "TABPES", "IDXPES01")
End Function
 
Public Function oTABPPC(Optional cSql As Variant) As csRS
Set oTABPPC = Abre(cSql, "TABPPC", "IDXPPC01")
End Function
 
Public Function oTABRAT(Optional cSql As Variant) As csRS
Set oTABRAT = Abre(cSql, "TABRAT", "IDXRAT01")
End Function
 
Public Function oTABREP(Optional cSql As Variant) As csRS
Set oTABREP = Abre(cSql, "TABREP", "IDXREP01")
End Function
 
Public Function oTABTRP(Optional cSql As Variant) As csRS
Set oTABTRP = Abre(cSql, "TABTRP", "IDXTRP01")
End Function
 
Public Function oTABUNF(Optional cSql As Variant) As csRS
Set oTABUNF = Abre(cSql, "TABUNF", "IDXUNF01")
End Function
 
Private Function Abre(Optional cSql As Variant, Optional cTabela As Variant, Optional cIndex As Variant) As csRS
Dim EnderecoBD As String

Open "rede.ini" For Input As #1
Line Input #1, EnderecoBD
Close #1

If IsMissing(cSql) Then
   Set Abre = csRSOpen(EnderecoBD & cteDBPES, (cTabela), cIndex)
 Else
   Set Abre = csRSOpen(EnderecoBD & cteDBPES, (cSql))
End If
End Function
