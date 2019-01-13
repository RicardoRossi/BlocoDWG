Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel
Public Class PlanilhaExcel

    Public Shared Function GetDadosColetor() As String

        Dim appXL As Excel.Application = New Excel.Application
        Dim wbXL As Excel.Workbook = appXL.Workbooks.Open($"{My.Application.Info.DirectoryPath}\nome dos arquivos.xlsx")
        Dim shXL As Excel.Worksheet = wbXL.Sheets(1)
        Dim raXL As Excel.Range = shXL.UsedRange
        shXL = wbXL.ActiveSheet

        Dim rowCount = raXL.Rows.Count
        Dim colCount = raXL.Columns.Count
        Dim coluna_A = ""
        Dim nomeArquivos = ""
        Try
            For i = 2 To rowCount
                Dim cellValue = CType(raXL.Cells(i, 1), Excel.Range) 'Cells retorna oject que e convertido para Range
                coluna_A = cellValue.Value.ToString
                nomeArquivos += coluna_A & Environment.NewLine
            Next
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Marshal.ReleaseComObject(raXL)
            Marshal.ReleaseComObject(shXL)
            wbXL.Close()
            Marshal.ReleaseComObject(wbXL)
            appXL.Quit()
            Marshal.ReleaseComObject(appXL)
        End Try
        Return nomeArquivos
    End Function
End Class
