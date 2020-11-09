Attribute VB_Name = "mDBPRM"
Option Explicit
 
Private Const cteDBPRM = "DBPRM.MDB"

Public Function oTABSRV(Optional cSql As Variant) As csRS
Set oTABSRV = Abre(cSql, "TABSRV", "IDXSRV01")
End Function
 
Public Function oTABDLR(Optional cSql As Variant) As csRS
Set oTABDLR = Abre(cSql, "TABDLR", "IDXDLR01")
End Function

Public Function oTABEMP(Optional cSql As Variant) As csRS
Set oTABEMP = Abre(cSql, "TABEMP")
End Function
 
Public Function oTABLOJ(Optional cSql As Variant) As csRS
Set oTABLOJ = Abre(cSql, "TABLOJ", "IDXLOJ01")
End Function
 
Public Function oTABOPF(Optional cSql As Variant) As csRS
Set oTABOPF = Abre(cSql, "TABOPF", "IDXOPF01")
End Function
 
Public Function oTABPST(Optional cSql As Variant) As csRS
Set oTABPST = Abre(cSql, "TABPST")
End Function
 
Public Function oTABPRM(Optional cSql As Variant) As csRS
Set oTABPRM = Abre(cSql, "TABPRM")
End Function

Public Function oTABTDP(Optional cSql As Variant) As csRS
Set oTABTDP = Abre(cSql, "TABTDP", "IDXTDP01")
End Function
 
Private Function Abre(Optional cSql As Variant, Optional cTabela As Variant, Optional cIndex As Variant) As csRS
Dim EnderecoBD As String

Open "rede.ini" For Input As #1
Line Input #1, EnderecoBD
Close #1

If IsMissing(cSql) Then
   Set Abre = csRSOpen(EnderecoBD & cteDBPRM, (cTabela), cIndex)
 Else
   Set Abre = csRSOpen(EnderecoBD & cteDBPRM, (cSql))
End If
End Function

