Attribute VB_Name = "RPRAT01M"
Option Explicit

Public Sub RPRAT01()
Dim TABRAT As csRS, oReport As New csReport

oReport.Prefix = AppSpoolPrefix
oReport.Path = "."
oReport.ReportName = "RPRAT01"
oReport.ReportTitle2 = "Listagem dos Ramos de Atividades"
oReport.ReportTitle1 = G_RL_Empresa
oReport.FinalBreak = True
If oReport.ShowDialogSets() Then
   HourGlassOn
   Set TABRAT = oTABRAT()
   
   If TABRAT.EOF Then
      iMsgBox "Nenhum registro processado ..."
   Else
      oReport.AddColumn "Código", 3, "000"
      oReport.AddColumn "Descrição", 20

      Do While Not TABRAT.EOF
         oReport.ColumnValue 1, TABRAT.RS!CODRAMATV, vbBlack
         oReport.ColumnValue 2, TABRAT.RS!DESRAMATV, vbBlack

         oReport.PrintDetail
         TABRAT.MoveNext
      Loop
      oReport.EndReport
   End If
   csRSClose TABRAT
   HourGlassOff
End If

Set oReport = Nothing
End Sub

