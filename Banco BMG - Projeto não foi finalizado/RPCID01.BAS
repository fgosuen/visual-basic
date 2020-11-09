Attribute VB_Name = "RPCID01M"
Option Explicit

Public Sub RPCID01()
Dim TABCID As csRS, TABUNF As csRS
Dim oReport As New csReport

oReport.Prefix = AppSpoolPrefix
oReport.Path = "."
oReport.ReportName = "RPCID01"
oReport.ReportTitle2 = "Listagem das Cidades"
oReport.ReportTitle1 = G_RL_Empresa
oReport.FinalBreak = True
If oReport.ShowDialogSets() Then
   HourGlassOn
   Set TABCID = oTABCID(): TABCID.Index = "IDXCID02"
   Set TABUNF = oTABUNF()
   
   If TABCID.EOF Then
      iMsgBox "Nenhum registro processado ..."
   Else
      oReport.AddColumn "Código", 5, "0000"
      oReport.AddColumn "Nome", 20
      oReport.AddColumn "Estado", 2
      oReport.AddColumn "", 20

      Do While Not TABCID.EOF
         oReport.ColumnValue 1, TABCID.RS!CODCID, vbBlack
         oReport.ColumnValue 2, TABCID.RS!NOMCID, vbBlack
         oReport.ColumnValue 3, TABCID.RS!undfed, vbBlack
         oReport.ColumnValue 4, RetornaNOMEST(TABUNF, TABCID.RS!undfed), vbBlack

         oReport.PrintDetail
         TABCID.MoveNext
      Loop
      oReport.EndReport
   End If
   csRSClose TABCID
   csRSClose TABUNF
   HourGlassOff
End If

Set oReport = Nothing
End Sub

Public Function RetornaNOMEST(TABUNF As csRS, sUndFed As Variant) As String
   TABUNF.RsSeek "=", sUndFed
   If Not TABUNF.Nomatch Then
      RetornaNOMEST = NullToSpc(TABUNF.RS!NOMUNF)
   Else
      RetornaNOMEST = "Não Cadastrado"
   End If
End Function


