Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel
Public Class PlanilhaExcel
    Shared appXL As Excel.Application = Nothing
    Public Shared Function GetDadosColetor() As List(Of Bloco)
        appXL = New Excel.Application 'CreateObject("Excel.Application")
        Dim wbXL As Excel.Workbook = Nothing
        Dim shXL As Excel.Worksheet = Nothing
        Dim raXL As Excel.Range = Nothing
        Dim workSheetCell As Excel.Range = Nothing
        Dim usedRange As Excel.Range = Nothing
        Dim cells As Excel.Range = Nothing
        Dim columns As Excel.Range = Nothing
        wbXL = appXL.Workbooks.Open($"{My.Application.Info.DirectoryPath}\..\nome dos arquivos.xlsx")
        shXL = wbXL.Sheets(1)
        raXL = shXL.UsedRange

        workSheetCell = shXL.Cells
        usedRange = shXL.UsedRange
        cells = usedRange.Cells
        columns = cells.Columns
        Dim rowCount = raXL.Rows.Count
        Dim colCount = raXL.Columns.Count

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
            'GC.Collect()
            'GC.WaitForPendingFinalizers()
            If wbXL IsNot Nothing Then
                If columns IsNot Nothing Then Marshal.ReleaseComObject(columns)
                If cells IsNot Nothing Then Marshal.ReleaseComObject(cells)
                If usedRange IsNot Nothing Then Marshal.ReleaseComObject(usedRange)
                If workSheetCell IsNot Nothing Then Marshal.ReleaseComObject(workSheetCell)
                If shXL IsNot Nothing Then Marshal.ReleaseComObject(shXL)
                wbXL.Close()
                'Marshal.FinalReleaseComObject(wbXL)
            End If
            'raXL = Nothing
            'shXL = Nothing
            'wbXL = Nothing
            appXL.Quit()
            'Marshal.FinalReleaseComObject(appXL)
        End Try
        Return listaDeBlocos
    End Function
    Public Shared Sub ExcelFinalizar()
        GC.Collect()
        GC.WaitForPendingFinalizers()
        'appXL = Nothing
        Marshal.FinalReleaseComObject(appXL)
    End Sub
End Class
