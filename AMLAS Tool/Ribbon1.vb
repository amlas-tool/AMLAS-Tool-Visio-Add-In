Imports Microsoft.Office.Tools.Ribbon

Public Class AMLAS
    Dim prevPage
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        'prevPage = Globals.ThisAddIn.Application.ActivePage.Name
        Button2.Label = "Back to " & vbCrLf & "previous page"
    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        prevPage = Globals.ThisAddIn.Application.ActivePage.Name
        Globals.ThisAddIn.Application.ActiveDocument.Pages(1).OpenDrawWindow()
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        Globals.ThisAddIn.Application.ActiveDocument.Pages.Item(prevPage).OpenDrawWindow()
    End Sub
End Class
