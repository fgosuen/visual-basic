Attribute VB_Name = "BRPPC01M"
Option Explicit

Public Function BRPPC01(TABPPC As RecordSet, codcli As Long) As Boolean
Dim Browse1 As New csBrowse, View1 As New csDbView

View1.RecordSet RSParent(TABPPC), TABPPC
View1.LowRange = Array(codcli)
View1.HighRange = Array(codcli)
Browse1.RecordSetBrowse = View1

Browse1.Caption = "Pessoas de Contato"
Browse1.BrowseCols = 2
Browse1.BrowseRows = 15
Browse1.Reposition = False
Browse1.SelectionEnabled = True

Browse1.AddColumn "NOMPES", 40, csLeft, "Nome", ""
Browse1.AddColumn "NUMTLF", 20, csLeft, "Telefone", ""

Browse1.ReSizeForm
Browse1.Show vbModal
BRPPC01 = Not View1.EOF

End Function
