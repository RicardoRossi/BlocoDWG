Imports System.IO
Imports SldWorks
Imports SwConst

Module ModuleMain
    Dim swApp As SldWorks.SldWorks
    Dim swModel As ModelDoc2
    Dim swCustPropMgr As CustomPropertyManager
    Dim swExt As ModelDocExtension
    Dim erro As Integer
    Dim aviso As Integer

    Sub Main()
        swApp = _swApp()
        Dim fullPathArquivoTemplate = My.Application.Info.DirectoryPath & "\..\Draw1.SLDDRW"
        'swApp.DocumentVisible(False, swDocumentTypes_e.swDocDRAWING)
        Dim desenho = swApp.OpenDoc6(fullPathArquivoTemplate, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_ReadOnly, "", erro, aviso)
        Dim tipoErroOpenDoc = CType(erro, swFileLoadError_e).ToString 'Nome do enumerador swFileLoadError_e dentro dele estão os tipos de erro
        If desenho Is Nothing Then
            MsgBox(tipoErroOpenDoc)
            Exit Sub
        End If

        swModel = swApp.ActiveDoc
        swExt = swModel.Extension
        swCustPropMgr = swExt.CustomPropertyManager("")

        Dim listaDeBlocos = PlanilhaExcel.GetDadosColetor()
        Dim listaArquivosRepetidos = ""
        For Each bloco In listaDeBlocos
            Dim pesoValor = bloco.peso
            Dim medidasValor As String = bloco.medidas
            Dim nomeArquivoValor = bloco.nomeArquivo

            Dim caminhoSalvarDWG = $"{My.Application.Info.DirectoryPath}\..\Arquivos DWG\{nomeArquivoValor}.dwg"
            If File.Exists(caminhoSalvarDWG) Then
                listaArquivosRepetidos += Path.GetFileName(caminhoSalvarDWG) & System.Environment.NewLine
                Continue For
            End If

            swCustPropMgr.Add3("nomeArquivo", swCustomInfoType_e.swCustomInfoText, nomeArquivoValor, swCustomPropertyAddOption_e.swCustomPropertyReplaceValue)
            swCustPropMgr.Add3("medidas", swCustomInfoType_e.swCustomInfoText, medidasValor, swCustomPropertyAddOption_e.swCustomPropertyReplaceValue)
            swCustPropMgr.Add3("peso", swCustomInfoType_e.swCustomInfoText, pesoValor, swCustomPropertyAddOption_e.swCustomPropertyReplaceValue)
            AlterarDimencoes(bloco, swModel, swExt)

            Dim salvouDoc As Boolean
            salvouDoc = swExt.SaveAs(caminhoSalvarDWG, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, erro, aviso)
            Dim tipoErroSaveAs4 = CType(erro, swFileSaveError_e).ToString 'Nome do enumerador swFileLoadError_e dentro dele estão os tipos de erro
            Dim tipoAvisoSaveAs4 = CType(erro, swFileSaveWarning_e).ToString
        Next
        MsgBox("Os arquivos já existem na pasta." & System.Environment.NewLine & listaArquivosRepetidos)
        swApp.ExitApp()
    End Sub

    Sub AlterarDimencoes(bloco As Bloco, swmodel As ModelDoc2, swext As ModelDocExtension)
        AlteraTamanhoDaFonteDaAnnotation(0.05)
        Dim boolstatus As Boolean
        boolstatus = swext.SelectByID2("COMPRIMENTO@Sketch5@Draw1.SLDDRW", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        Dim myDimension As Dimension
        myDimension = swmodel.Parameter("COMPRIMENTO@Sketch5")
        myDimension.SystemValue = bloco.comprimento / 100
        boolstatus = swmodel.Extension.SelectByID2("PROFUNDIDADE@Sketch5@Draw1.SLDDRW", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
        myDimension = swmodel.Parameter("PROFUNDIDADE@Sketch5")
        myDimension.SystemValue = bloco.profundidade / 100

        If bloco.comprimento > ((bloco.nomeArquivo.Length + 1) * 4) Then
            Dim fator = 4.0
            boolstatus = swmodel.Extension.SelectByID2("NOME_ARQUIVO@Sketch5@Draw1.SLDDRW", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = swmodel.Parameter("NOME_ARQUIVO@Sketch5")
            myDimension.SystemValue = CalculaComprimentoDoRetangulo(bloco.nomeArquivo, fator)
            boolstatus = swmodel.Extension.SelectByID2("DIMENSOES@Sketch5@Draw1.SLDDRW", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = swmodel.Parameter("DIMENSOES@Sketch5")
            myDimension.SystemValue = CalculaComprimentoDoRetangulo(bloco.medidas, fator)
            boolstatus = swmodel.Extension.SelectByID2("PESO@Sketch5@Draw1.SLDDRW", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = swmodel.Parameter("PESO@Sketch5")
            myDimension.SystemValue = CalculaComprimentoDoRetangulo(bloco.peso, fator)
            swmodel.ClearSelection2(True)
        Else
            AlteraTamanhoDaFonteDaAnnotation(0.035)
            boolstatus = swmodel.Extension.SelectByID2("NOME_ARQUIVO@Sketch5@Draw1.SLDDRW", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = swmodel.Parameter("NOME_ARQUIVO@Sketch5")
            myDimension.SystemValue = CalculaComprimentoDoRetangulo(bloco.nomeArquivo, 2.7)
            boolstatus = swmodel.Extension.SelectByID2("DIMENSOES@Sketch5@Draw1.SLDDRW", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = swmodel.Parameter("DIMENSOES@Sketch5")
            myDimension.SystemValue = CalculaComprimentoDoRetangulo(bloco.medidas, 2.7)
            boolstatus = swmodel.Extension.SelectByID2("PESO@Sketch5@Draw1.SLDDRW", "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
            myDimension = swmodel.Parameter("PESO@Sketch5")
            myDimension.SystemValue = CalculaComprimentoDoRetangulo(bloco.peso, 3.0)
            swmodel.ClearSelection2(True)
        End If
    End Sub

    Function CalculaComprimentoDoRetangulo(campo As String, fator As Double) As Double
        Dim comprimentoDoRetangulo As Double
        comprimentoDoRetangulo = (campo.Length * fator) / 100 'Cada caracter é multiplicado pelo fator e depois convertido metro so SolidWorks para centimetro.
        Return comprimentoDoRetangulo
    End Function

    Sub AlteraTamanhoDaFonteDaAnnotation(tamanhoFonte As Double)
        Dim boolstatus As Boolean
        Dim myTextFormat As Object
        Dim swApp As SldWorks.SldWorks
        Dim swModel As ModelDoc2
        Dim swExt As ModelDocExtension
        swApp = _swApp()
        swModel = swApp.ActiveDoc
        swExt = swModel.Extension
        myTextFormat = swExt.GetUserPreferenceTextFormat(swUserPreferenceTextFormat_e.swDetailingNoteTextFormat, 0)
        myTextFormat.CharHeight = tamanhoFonte
        boolstatus = swExt.SetUserPreferenceTextFormat(swUserPreferenceTextFormat_e.swDetailingNoteTextFormat, 0, myTextFormat)
    End Sub
End Module
