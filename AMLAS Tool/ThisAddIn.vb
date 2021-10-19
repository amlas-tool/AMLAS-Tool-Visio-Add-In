Imports Microsoft.Office.Interop.Visio



Public Class ThisAddIn
    'global consts for colours
    Dim stage_colour(7) As String
    Dim stages(7) As String



    Private Sub ThisAddIn_Startup() Handles Me.Startup

        stages = {"", "Stage 1", "Stage 2", "Stage 3", "Stage 4", "Stage 5", "Stage 6"} 'amlas overview index is zero
        'colours need to be formatted for shapesheet entries, not RGB
        stage_colour(1) = "THEMEGUARD(MSOTINT(THEMEVAL(""AccentColor6""),60))"
        stage_colour(2) = "THEMEGUARD(MSOTINT(THEMEVAL(""AccentColor5""),60))"
        stage_colour(3) = "THEMEGUARD(MSOTINT(THEMEVAL(""AccentColor4""),60))"
        stage_colour(4) = "THEMEGUARD(MSOTINT(THEMEVAL(""AccentColor""),60))"
        stage_colour(5) = "THEMEGUARD(MSOTINT(RGB(255,255,255),-15))"
        stage_colour(6) = "THEMEGUARD(MSOTINT(THEMEVAL(""AccentColor2""),60))"


        'disable developer mode ?
        'Application.Settings.DeveloperMode
        'If Application.Settings.DeveloperMode Then Application.Settings.DeveloperMode = Not Application.Settings.DeveloperMode

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub Application_ShapeAdded(Shape As Shape) Handles Application.ShapeAdded
        'on drop, add stage colour to shape and prompt user to enter shape data
        Dim vsoMaster As Visio.Master
        Dim visioDocs As Visio.Documents = Me.Application.Documents
        Dim activePage As String

        activePage = Globals.ThisAddIn.Application.ActivePage.Name
        'Get the Master property of the shape. 
        vsoMaster = Shape.Master

        'Check whether the shape has a master. If not, 
        'the shape was created locally. 
        If Not vsoMaster Is Nothing Then
            'create background colour for dropped shapes
            For i = 1 To stages.Length - 1
                If activePage.Contains(stages(i)) Then
                    If Not vsoMaster.Name.Contains("Document") Then 'documents have no colour
                        If vsoMaster.Name.Contains("Justification") Or vsoMaster.Name.Contains("Assumption") Then
                            For Each shapewithin As Visio.Shape In Shape.Shapes
                                shapewithin.CellsU("Fillforegnd").FormulaU = stage_colour(i)
                            Next shapewithin
                        Else
                            Shape.CellsU("Fillforegnd").FormulaU = stage_colour(i)
                        End If
                    End If
                End If

            Next i

        End If



    End Sub

    Private Sub Application_ShapeChanged(Shape As Shape) Handles Application.ShapeChanged
        'Dim message As String
        Dim currentpage As String
        Dim correspondingArgPage As String = ""
        Dim page As Visio.Page

        'create context referenced from this shape in corresponding stage argument
        If Shape.Name.Contains("Document") Then


            currentpage = Globals.ThisAddIn.Application.ActivePage.Name
                For i = 1 To stages.Length - 1      'skip overview page i=0
                    If currentpage.Contains(stages(i)) Then

                        For Each page In Globals.ThisAddIn.Application.ActiveDocument.Pages
                            If page.Name <> currentpage And page.Name.Contains(stages(i)) Then
                                correspondingArgPage = page.Name
                            End If
                        Next page

                    End If
                Next i

            'message = Shape.Name & " would become a referenced Context " & Shape.CellsU("Prop.ArtID").FormulaU & " added to:" & vbCrLf & correspondingArgPage & "." & vbCrLf & vbCrLf & "Linkage in the argument for this Context must be done manually."
            'MsgBox(message)

        ElseIf Shape.Name = "Dynamic connector" Then 'format arrow if user adds automatic shape from activity
            Shape.CellsU("LinePattern").FormulaU = "23"
            Shape.CellsU("EndArrow").FormulaU = "13"

        End If

    End Sub
End Class


