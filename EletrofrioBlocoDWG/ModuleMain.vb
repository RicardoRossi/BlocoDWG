Imports System.IO
Imports SldWorks
Imports SwConst
Imports System.Console

Module ModuleMain
    Dim swApp As SldWorks.SldWorks
    Dim swModel As ModelDoc2

    Sub Main()
        swApp = _swApp()
        swApp.SendMsgToUser("conectado")

    End Sub
End Module
