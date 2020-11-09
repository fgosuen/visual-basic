Attribute VB_Name = "BRLOJ01M"
Option Explicit
 
Public Function BRLOJ01(TABLOJ As RecordSet) As Boolean
Dim Browse1 As New csBrowse
 
Browse1.Caption = "LOJAS"
Browse1.BrowseCols = 2
Browse1.BrowseRows = 15
 
Browse1.RecordSetBrowse = TABLOJ
Browse1.AddColumn "codloj", 6, 0, "Código", "0000"
Browse1.AddColumn "nomloj", 30, 0, "Loja"
Browse1.ReSizeForm
 
Browse1.Show vbModal
BRLOJ01 = Not TABLOJ.EOF
End Function
 
