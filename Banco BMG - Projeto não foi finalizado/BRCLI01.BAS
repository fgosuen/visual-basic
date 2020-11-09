Attribute VB_Name = "BRCLI01M"
Option Explicit

Public Function BRCLI01(TABCLI As RecordSet) As Boolean
Dim Browse1 As New csBrowse

Browse1.RecordSetBrowse = TABCLI
Browse1.RecordSetIndex = "IDXCLI02"

Browse1.Caption = "Clientes"
Browse1.BrowseCols = 2
Browse1.BrowseRows = 15
Browse1.Reposition = True
Browse1.SelectionEnabled = True

Browse1.AddColumn "CODCLI", 6, csLeft, "Código", "00000"
Browse1.AddColumn "NOMCLI", 40, csLeft, "Nome", ""

Browse1.ReSizeForm
Browse1.Show vbModal
BRCLI01 = Not TABCLI.EOF

End Function
