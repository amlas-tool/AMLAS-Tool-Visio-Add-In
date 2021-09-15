Imports Microsoft.Office.Tools.Ribbon

Public Class AMLAS
    Dim prevPage
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        'prevPage = Globals.ThisAddIn.Application.ActivePage.Name
        Button2.Label = "Back to " & vbCrLf & "previous page"
        Button3.Label = "Create new doc" & vbCrLf & "from AMLAS Template"
    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        prevPage = AMLAS_Tool.Globals.ThisAddIn.Application.ActivePage.Name
        AMLAS_Tool.Globals.ThisAddIn.Application.ActiveWindow.Page = "AMLAS Process Overview"
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        AMLAS_Tool.Globals.ThisAddIn.Application.ActiveWindow.Page = prevPage
    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        Dim docPath As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) + "\Visio Add In\AMLAS Tool.vstx"
        AMLAS_Tool.Globals.ThisAddIn.Application.Documents.Add(docPath)
    End Sub
End Class
