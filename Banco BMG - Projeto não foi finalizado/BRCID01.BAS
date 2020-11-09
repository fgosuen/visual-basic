Attribute VB_Name = "BRCID01M"
Option Explicit
 
Public Function BRCID01(TABCID As RecordSet) As Boolean
Dim Browse1 As New csBrowse
 
Browse1.Caption = "Cidades"
Browse1.BrowseCols = 3
Browse1.BrowseRows = 15
Browse1.Reposition = True
 
Browse1.RecordSetBrowse = TABCID
Browse1.RecordSetIndex = "IDXCID02"
Browse1.AddColumn "codcid", 6, csCenter, "Código", "0000"
Browse1.AddColumn "nomcid", 30, csLeft, "Nome"
Browse1.AddColumn "undfed", 3, csCenter, "UF"
Browse1.ReSizeForm
 
Browse1.Show vbModal
BRCID01 = Not TABCID.EOF
End Function
 

