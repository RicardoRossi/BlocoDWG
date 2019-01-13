Module SolidWorksSingleton
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Function _swApp() As SldWorks.SldWorks
        Try
            If swApp Is Nothing Then
                swApp = GetObject("", "SldWorks.Application")
                swApp.Visible = True
                Return swApp
            Else
                Return swApp
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function

    Function _swModel() As SldWorks.ModelDoc2
        swModel = _swApp.ActiveDoc
        Return swModel
    End Function

End Module
