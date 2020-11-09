Attribute VB_Name = "BRPES01M"
Option Explicit

Public Function BRPES01(TABPES As RecordSet) As Boolean
Dim Browse1 As New csBrowse

Browse1.RecordSetBrowse = TABPES

Browse1.Caption = "Cadastro de Pessoas"
Browse1.BrowseCols = 4
Browse1.BrowseRows = 15
Browse1.Reposition = True
Browse1.SelectionEnabled = True

Browse1.RecordSetIndex = "IDXPES03"
Browse1.AddColumn "CODPES", 6, csLeft, "Código", "000000"
Browse1.AddColumn "NOMRES", 15, csLeft, "Nome Resumido", ""
Browse1.AddColumn "NOMPES", 40, csLeft, "Nome", ""
Browse1.AddColumn "NUMTLF", 20, csLeft, "Telefone", ""

Browse1.ReSizeForm
Browse1.Show vbModal
BRPES01 = Not TABPES.EOF

End Function

Public Function BRPES02(TABPES As RecordSet) As Boolean
Dim Browse1 As New csBrowse

Browse1.RecordSetBrowse = TABPES

Browse1.Caption = "Cadastro de Pessoas"
Browse1.BrowseCols = 4
Browse1.BrowseRows = 15
Browse1.Reposition = True
Browse1.SelectionEnabled = True

Browse1.AddColumn "CODPES", 6, csLeft, "Código", "000000"
Browse1.AddColumn "NOMRES", 15, csLeft, "Nome Resumido", ""
Browse1.AddColumn "NOMPES", 40, csLeft, "Nome", ""
Browse1.AddColumn "NUMTLF", 20, csLeft, "Telefone", ""

Browse1.ReSizeForm
Browse1.Show vbModal
BRPES02 = Not TABPES.EOF

End Function

