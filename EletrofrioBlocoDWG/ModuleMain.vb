Imports System.IO
Imports SldWorks
Imports SwConst
Imports System.Console
Imports System.Deployment

Module ModuleMain
    Dim swApp As SldWorks.SldWorks
    Dim swModel As ModelDoc2
    Dim swCustPropMgr As CustomPropertyManager
    Dim swExt As ModelDocExtension
    Dim erro As String
    Dim aviso As String

    Sub Main()
        Dim codigos As List(Of String) = New List(Of String) From {"MONTCPRP0001"}

        swApp = _swApp()
        'swApp.SendMsgToUser(My.Application.Info.DirectoryPath)
        Dim fullPathArquivoTemplate = My.Application.Info.DirectoryPath & "\Draw1.SLDDRW"
        Dim desenho = swApp.OpenDoc6(fullPathArquivoTemplate, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_ReadOnly, "", erro, aviso)

        Dim tipoErro = CType(erro, swFileLoadError_e).ToString
        If desenho Is Nothing Then
            MsgBox(tipoErro)
            Exit Sub
        End If

        swModel = swApp.ActiveDoc
        swExt = swModel.Extension
        swCustPropMgr = swExt.CustomPropertyManager("")

        For Each codigo In codigos
            Dim pesoValor As String
            Dim medidasValor As String
            Dim nomeArquivoValor = codigo

            swCustPropMgr.Add3("nomeArquivo", swCustomInfoType_e.swCustomInfoText, nomeArquivoValor, swCustomPropertyAddOption_e.swCustomPropertyReplaceValue)
            swCustPropMgr.Add3("medidas", swCustomInfoType_e.swCustomInfoText, "123x456x789", swCustomPropertyAddOption_e.swCustomPropertyReplaceValue)
            swCustPropMgr.Add3("peso", swCustomInfoType_e.swCustomInfoText, "11111" & "kg", swCustomPropertyAddOption_e.swCustomPropertyReplaceValue)
            swModel.SaveAs($"{My.Application.Info.DirectoryPath}\{nomeArquivoValor}.dwg")
        Next

        Dim lista = PlanilhaExcel.GetDadosColetor()
        Write(lista)
        ReadKey()
        'swModel.EditRebuild3()
    End Sub
End Module
