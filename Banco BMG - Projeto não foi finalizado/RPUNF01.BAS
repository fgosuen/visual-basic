Attribute VB_Name = "RPUNF01M"
Option Explicit

Public Sub RPUNF01()
Dim TABUNF As csRS, oReport As New csReport

oReport.Prefix = AppSpoolPrefix
oReport.Path = "."
oReport.ReportName = "RPUNF01"
oReport.ReportTitle2 = "Listagem dos Estados"
oReport.ReportTitle1 = G_RL_Empresa
oReport.FinalBreak = True
If oReport.ShowDialogSets() Then
   HourGlassOn
   Set TABUNF = oTABUNF()
   
   If TABUNF.EOF Then
      iMsgBox "Nenhum registro processado ..."
   Else
      oReport.AddColumn "Sigla", 2
      oReport.AddColumn "Nome", 20
      oReport.AddColumn "%Crédito ICMS", 10, csFmMoney, csRight
      oReport.AddColumn "%ICMS Consumo", 10, csFmMoney, csRight
      oReport.AddColumn "%ICMS Revenda", 10, csFmMoney, csRight

      Do While Not TABUNF.EOF
         oReport.ColumnValue 1, TABUNF.RS!undfed, vbBlack
         oReport.ColumnValue 2, TABUNF.RS!NOMUNF, vbBlack
         oReport.ColumnValue 3, TABUNF.RS!CREICM, vbBlack
         oReport.ColumnValue 4, TABUNF.RS!DEBICMCON, vbBlack
         oReport.ColumnValue 5, TABUNF.RS!DEBICMRVD, vbBlack

         oReport.PrintDetail
         TABUNF.MoveNext
      Loop
      oReport.EndReport
   End If
   csRSClose TABUNF
   HourGlassOff
End If

Set oReport = Nothing
End Sub


