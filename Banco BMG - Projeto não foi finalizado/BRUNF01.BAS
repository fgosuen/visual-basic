Attribute VB_Name = "BRUNF01M"
Option Explicit
 
Public Function BRUNF01(TABUNF As RecordSet) As Boolean
Dim Browse1 As New csBrowse
 
Browse1.Caption = "Estados"
Browse1.BrowseCols = 2
Browse1.BrowseRows = 15
 
Browse1.RecordSetBrowse = TABUNF
Browse1.AddColumn "undfed", 3, csLeft, "UF"
Browse1.AddColumn "nomunf", 25, csLeft, "Estado"
Browse1.ReSizeForm
 
Browse1.Show vbModal
BRUNF01 = Not TABUNF.EOF
End Function
 
 

