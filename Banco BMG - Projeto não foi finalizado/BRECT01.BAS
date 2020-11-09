Attribute VB_Name = "BRECT01M"
Option Explicit
 
Public Function BRECT01(TABECT As RecordSet, codtrp As Long) As Boolean
Dim Browse1 As New csBrowse, View1 As New csDbView, TABCID As csRS
 
Set TABCID = oTABCID()
 
View1.RecordSet RSParent(TABECT), TABECT
View1.LowRange = Array(codtrp)
View1.HighRange = Array(codtrp)
View1.AddRelation TABCID.RS, Array("codcid")
Browse1.RecordSetBrowse = View1
 
Browse1.Caption = "Endereços de Coleta"
Browse1.BrowseCols = 4
Browse1.BrowseRows = 15
Browse1.Reposition = False
Browse1.SelectionEnabled = True
 
Browse1.AddColumn "1:NOMCID", 15, csLeft, "Cidade", ""
Browse1.AddColumn "DESEND", 30, csLeft, "Endereço", ""
Browse1.AddColumn "NOMBAI", 15, csLeft, "Bairro", ""
Browse1.AddColumn "NUMTLF", 15, csLeft, "Telefone", ""
 
Browse1.ReSizeForm
Browse1.Show vbModal
BRECT01 = Not View1.EOF
 
End Function

