Attribute VB_Name = "BRRAT01M"
Option Explicit
 
Public Function BRRAT01(TABRAT As RecordSet) As Boolean
Dim Browse1 As New csBrowse
 
Browse1.RecordSetBrowse = TABRAT
Browse1.RecordSetIndex = "IDXRAT02"
 
Browse1.Caption = "Ramos de Atividades"
Browse1.BrowseCols = 2
Browse1.BrowseRows = 15
Browse1.Reposition = True
Browse1.SelectionEnabled = True
 
Browse1.AddColumn "CODRAMATV", 6, csCenter, "Código", "000"
Browse1.AddColumn "DESRAMATV", 30, csLeft, "Descrição", ""
 
Browse1.ReSizeForm
Browse1.Show vbModal
BRRAT01 = Not TABRAT.EOF
 
End Function

