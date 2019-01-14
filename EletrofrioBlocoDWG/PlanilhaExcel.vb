Imports Excel = Microsoft.Office.Interop.Excel
Public Class PlanilhaExcel
    Shared appXL As Excel.Application = New Excel.Application
    Shared wbXL As Excel.Workbook = appXL.Workbooks.Open($"{My.Application.Info.DirectoryPath}\..\nome dos arquivos.xlsx")
    Shared shXL As Excel.Worksheet = wbXL.Sheets(1)
    Shared raXL As Excel.Range = shXL.UsedRange
    Shared rowCount = raXL.Rows.Count
    Shared colCount = raXL.Columns.Count
    Public Shared Function GetDadosColetor() As List(Of Bloco)
        shXL = wbXL.ActiveSheet
        Dim listaDeBlocos As New List(Of Bloco)
        Try
            For i = 2 To rowCount
                Dim bloco As New Bloco
                Dim cellValue = CType(raXL.Cells(i, 1), Excel.Range) 'Cells retorna oject que e convertido para Range
                bloco.nomeArquivo = cellValue.Value.ToString
                cellValue = CType(raXL.Cells(i, 2), Excel.Range) 'Cells retorna oject que e convertido para Range
                bloco.comprimento = cellValue.Value.ToString
                cellValue = CType(raXL.Cells(i, 3), Excel.Range) 'Cells retorna oject que e convertido para Range
                bloco.profundidade = cellValue.Value.ToString
                cellValue = CType(raXL.Cells(i, 4), Excel.Range) 'Cells retorna oject que e convertido para Range
                bloco.altura = cellValue.Value.ToString
                cellValue = CType(raXL.Cells(i, 5), Excel.Range) 'Cells retorna oject que e convertido para Range
                bloco.peso = cellValue.Value.ToString & "kg"
                bloco.medidas = $"{bloco.comprimento}x{bloco.profundidade}x{bloco.altura}"
                listaDeBlocos.Add(bloco)
            Next
        Finally
            'wbXL.Save()
            raXL = Nothing
            shXL = Nothing
            wbXL = Nothing
            appXL.Quit()
            appXL = Nothing
        End Try
        Return listaDeBlocos
    End Function
End Class
