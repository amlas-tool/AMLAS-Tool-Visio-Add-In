Imports Microsoft.Office.Tools.Ribbon

Public Class AMLAS
    Dim prevPage
    Dim stages(7) As String

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        stages = {"", "Stage 1", "Stage 2", "Stage 3", "Stage 4", "Stage 5", "Stage 6"}
        Button2.Label = "Back to " & vbCrLf & "previous page"
        Button3.Label = "Create"
        Button4.Label = "Toggle between" & vbCrLf & "process && arg"
        Group1.Label = "Create new doc" & vbCrLf & "from AMLAS Template"

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Try
            prevPage = Globals.ThisAddIn.Application.ActivePage.Name
            Globals.ThisAddIn.Application.ActiveWindow.Page = "AMLAS Process Overview"
        Catch ex As System.NullReferenceException
            MsgBox("Please load a document or click Create " & vbCrLf & "button to create new from template.", MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        Try
            Globals.ThisAddIn.Application.ActiveWindow.Page = prevPage
        Catch ex As System.NullReferenceException
            MsgBox("Please load a document or click Create " & vbCrLf & "button to create new from template.", MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        Dim docPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\Visio Add In\AMLAS Tool.vstx"
        Globals.ThisAddIn.Application.Documents.Add(docPath)
        Globals.ThisAddIn.Application.ActiveWindow.Page = "AMLAS Process Overview"
    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click
        Dim currentpage As String
        Dim page As Visio.Page
        Try


            currentpage = Globals.ThisAddIn.Application.ActivePage.Name
            For i = 1 To stages.Length - 1
                If currentpage.Contains(stages(i)) Then

                    For Each page In Globals.ThisAddIn.Application.ActiveDocument.Pages
                        If page.Name <> currentpage And page.Name.Contains(stages(i)) Then
                            Globals.ThisAddIn.Application.ActiveWindow.Page = page.Name
                        End If
                    Next page


                End If
            Next i

        Catch ex As System.NullReferenceException
            MsgBox("Please load a document or click Create " & vbCrLf & "button to create new from template.", MsgBoxStyle.Exclamation)
        End Try


    End Sub
End Class
