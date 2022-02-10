Imports System
Imports System.Diagnostics

Friend Class HandleMouseEvents
    '// <summary>This function finds the shapes that are
    '// at the specified location.  This is usually used in response to
    '// a mouse click event.</summary>
    '// <param name="clickedPage">Page that contains the shapes</param>
    '// <param name="clickedLocationX">X coordinate in Visio units
    '// </param>
    '// <param name="clickedLocationY">Y coordinate in Visio units
    '// </param>
    '// <param name="tolerance">Distance from the location that a shape
    '// can be to be considered clicked</param>
    '// <returns>Selection object containing the list of shapes that are
    '// at the specified location (The Selection object is Nothing if an
    '// error occurs.)</returns>        
    Public Shared Function getMouseClickShapes(
            ByVal clickedPage As Microsoft.Office.Interop.Visio.Page,
            ByVal clickedLocationX As Double,
            ByVal clickedLocationY As Double,
            ByVal tolerance As Double) _
            As Microsoft.Office.Interop.Visio.Selection

        Dim clickedShapes As Microsoft.Office.Interop.Visio.Selection = Nothing

        Try

            ' Use the SpatialSearch method of the page to get the list
            ' of shapes at the location.
            clickedShapes = clickedPage.SpatialSearch(
                    clickedLocationX, clickedLocationY,
                    CInt(Microsoft.Office.Interop.Visio.
                    VisSpatialRelationCodes.visSpatialContainedIn),
                    tolerance,
                    CInt(Microsoft.Office.Interop.Visio.
                    VisSpatialRelationFlags.visSpatialFrontToBack))

        Catch errorThrown As System.Runtime.InteropServices.COMException
            Debug.WriteLine(errorThrown.Message)
        End Try

        Return clickedShapes

    End Function

    Friend Shared Sub ReplaceShapeWithRectangle(cShape As Visio.Shape, page As Visio.Page, newName As String, text As String)
        Dim currentVal1 As Integer = cShape.CellsU("LockDelete").ResultIU
        Dim currentVal2 As Integer = cShape.CellsU("LockGroup").ResultIU
        Dim currentVal3 As Integer = cShape.CellsU("LockFromGroupFormat").ResultIU

        cShape.CellsU("LockDelete").ResultIU = 0
        cShape.CellsU("LockGroup").ResultIU = 0 'shp.Protection.LockBegin.Value = True
        cShape.CellsU("LockFromGroupFormat").ResultIU = 0

        Dim cellX As Double = cShape.Cells("PinX").ResultIU
        Dim cellY As Double = cShape.Cells("PinY").ResultIU

        Dim selToDel As Visio.Selection
        selToDel = page.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty)
        Call selToDel.Select(cShape, Visio.VisSelectArgs.visSelect)
        selToDel.Delete()
        'Globals.ThisAddIn.Application.ActivePage.Drop(srcShape, cellX, cellY)
        cShape = page.DrawRectangle(cellX - 0.125, cellY - 0.125, cellX + 0.125, cellY + 0.125)
        cShape.Name = newName
        If text IsNot Nothing Then
            cShape.Text = text
        End If
        Call selToDel.Select(cShape, Visio.VisSelectArgs.visDeselectAll)
        cShape.CellsU("LockDelete").ResultIU = currentVal1
        cShape.CellsU("LockGroup").ResultIU = currentVal2
        cShape.CellsU("LockFromGroupFormat").ResultIU = currentVal3
    End Sub

    'Detect when the user clicks on somehting - mouseDown event
    'Calls diiferent functionality depending where they clicked
    Friend Shared Sub HandleMouseDown(ByVal mouseButton As Integer,
            ByVal buttonState As Integer,
            ByVal clickLocationX As Double,
            ByVal clickLocationY As Double,
            ByRef cancelProcessing As Boolean,
            ByRef clickedWindow As Microsoft.Office.Interop.Visio.Window)

        Dim locX, locY As String
        locX = CStr(clickLocationX)
        locY = CStr(clickLocationY)
        'MsgBox("mouse clicked at " + loc, MsgBoxStyle.Information)
        Debug.Print("x is: " + locX)
        Debug.Print("y is: " + locY)
        Const tolerance As Double = 0.0001

        Dim clickedShapes As Microsoft.Office.Interop.Visio.Selection = Nothing
        Dim otherPageShapes As Microsoft.Office.Interop.Visio.Selection = Nothing
        Dim nextShape As Integer
        Dim message As New System.Text.StringBuilder()

        Dim isCircLeft As Boolean = False
        Dim circLeftShape As Visio.Shape = Nothing
        Dim isCircRight As Boolean = False
        Dim circRightShape As Visio.Shape = Nothing
        Dim isDiamond As Boolean = False
        Dim isUpdateOverview As Boolean = False
        Dim otherPageShape As Visio.Shape = Nothing

        Try
            'Store window where mouse click occurred
            clickedWindow = Globals.ThisAddIn.Application.ActiveWindow
            ' Check if the left mouse mouseButton caused this event
            ' to occur.
            If mouseButton = CInt(Microsoft.Office.Interop.Visio.
                    VisKeyButtonFlags.visMouseLeft) Then

                ' Get the list of shapes at the click location.
                clickedShapes = getMouseClickShapes(
                        clickedWindow.PageAsObj, clickLocationX,
                        clickLocationY, tolerance)


                ' Check if any shapes were found.
                If ((Not clickedShapes Is Nothing) And
                        (clickedShapes.Count > 0)) Then

                    'message.Append(shapesFoundPrompt & System.Environment.NewLine)
                    Dim targetShape As Visio.Shape = Nothing

                    ' Loop through the collection of clicked shapes.
                    For nextShape = 1 To clickedShapes.Count

                        'If user clicked circle to set number of performance SRs
                        If (clickedShapes.Item(nextShape).Name.Trim().StartsWith("CircleLeft") And Globals.ThisAddIn.Application.ActivePage.NameU = ThisAddIn.stageNames(1)) Then
                            isCircLeft = True
                            circLeftShape = clickedShapes.Item(nextShape)
                            Dim otherPage As Visio.Page
                            'Dim visioPages As Visio.Pages = Globals.ThisAddIn.Application.ActiveDocument.Pages

                            otherPage = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(6))
                            otherPageShape = ThisAddIn.GetShape(otherPage, circLeftShape)
                            NumSRs.ReadSRsFromFile()
                        End If
                        'If user clicked circle to set number of robustness SRs
                        If (clickedShapes.Item(nextShape).Name.Trim().StartsWith("CircleRight") And Globals.ThisAddIn.Application.ActivePage.NameU = ThisAddIn.stageNames(1)) Then
                            isCircRight = True
                            circRightShape = clickedShapes.Item(nextShape)
                            Dim otherPage As Visio.Page
                            'Dim visioPages As Visio.Pages = Globals.ThisAddIn.Application.ActiveDocument.Pages

                            otherPage = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(6))
                            otherPageShape = ThisAddIn.GetShape(otherPage, circRightShape)
                            NumSRs.ReadSRsFromFile()
                        End If
                        'If user clicked diamond to choose branches to remove 
                        If (clickedShapes.Item(nextShape).Name.Trim().StartsWith("Filled diamond") And Globals.ThisAddIn.Application.ActivePage.NameU = ThisAddIn.stageNames(4)) Then
                            targetShape = clickedShapes.Item(nextShape)
                            isDiamond = True
                            NumSRs.ReadSRsFromFile()
                        End If
                        If (clickedShapes.Item(nextShape).Name.Trim().StartsWith("Refresh") And Globals.ThisAddIn.Application.ActivePage.NameU = ThisAddIn.stageNames(6)) Then
                            isUpdateOverview = True
                            NumSRs.ReadSRsFromFile()
                        End If
                    Next nextShape
                    'If the user clicked the left circle then process their responses
                    'We will create an overview page and set the number of performence SRs
                    If ((clickedWindow.Application.AlertResponse = 0) And isCircLeft) Then
                        'Ask the user how many performance SRs (below the left circle) they want 
                        Dim getMessage, Title, DefaultVal, MyValue
                        getMessage = "How many performance SRs are there? [1-6]"    ' Set prompt.
                        Title = "How many SRs are there?"    ' Set title.
                        DefaultVal = "1"    ' Set default.
                        ' Display message, title, and default value.
                        MyValue = InputBox(getMessage, Title, DefaultVal)
                        If ((Not MyValue Is Nothing) And (Not MyValue = "")) Then
                            Dim num As Integer
                            'If TypeOf MyValue Is Integer Then
                            Try
                                num = CInt(Int(Convert.ToString(MyValue)))

                                If (num > 0 And num < 7) Then
                                    Try
                                        CreateOverview.Create_Overview(num, True)
                                        NumSRs.SetNumPerformanceSRsToFile(num)
                                    Catch ex As Exception
                                        MsgBox(ex.Message)
                                    End Try
                                Else
                                    If (Not num = 0) Then
                                        MsgBox("Please enter a number between 0 to 6.")
                                    End If
                                End If
                                If (num >= 0 And num < 7) Then
                                    'Swap filled left circle with empty square
                                    'Dim visioDocs As Visio.Documents = Globals.ThisAddIn.Application.Documents
                                    'Dim rep As Object = visioDocs("AMLAS Tool Stencil.vssx").Masters("Clear Diamond")
                                    'targetShape.ReplaceShape(rep)
                                    'Dim page As String = 
                                    ReplaceShapeWithRectangle(circLeftShape, Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(1)), "CircleToSquareLeft", NumSRs.GetNumPerformanceSRs.ToString)
                                    CreateOverview.Update_Overview()
                                    MsgBox("We have added new instances to Stage 5 Argument Patterns" + vbCrLf + "and updated the Multi-Overview to reflect the additions.")
                                End If
                            Catch ex As Exception
                                MsgBox("Please enter a number between 0 to 6.")
                            End Try

                        End If
                        NumSRs.WriteSRsToFile()
                    End If
                    'If the user clicked the right circle then process their responses
                    'We will create an overview page and set the number of robustness SRs
                    If ((clickedWindow.Application.AlertResponse = 0) And isCircRight) Then
                        'Ask the user how many robustness SRs (below the right circle) they want 
                        Dim getMessage, Title, DefaultVal, MyValue
                        getMessage = "How many robustness SRs are there? [0-6]"    ' Set prompt.
                        Title = "How many SRs are there?"    ' Set title.
                        DefaultVal = "1"    ' Set default.
                        ' Display message, title, and default value.
                        MyValue = InputBox(getMessage, Title, DefaultVal)
                        If ((Not MyValue Is Nothing) And (Not MyValue = "")) Then
                            Dim num As Integer
                            'If TypeOf MyValue Is Integer Then
                            Try
                                num = CInt(Int(Convert.ToString(MyValue)))


                                If (num > 0 And num < 7) Then
                                    Try
                                        CreateOverview.Create_Overview(num, False)
                                        NumSRs.SetNumRobustnessSRsToFile(num)
                                    Catch ex As Exception
                                        MsgBox(ex.Message)
                                    End Try
                                Else
                                    If (Not num = 0) Then
                                        MsgBox("Please enter a number between 0 to 6.")
                                    End If
                                End If
                                If (num >= 0 And num < 7) Then
                                    'Swap filled right circle with empty square
                                    'Dim visioDocs As Visio.Documents = Globals.ThisAddIn.Application.Documents
                                    'Dim rep As Object = visioDocs("AMLAS Tool Stencil.vssx").Masters("Clear Diamond")
                                    'targetShape.ReplaceShape(rep)
                                    'Dim page As String = "Stage 2: Assurance Argument Pattern for ML Safety Requirements"
                                    ReplaceShapeWithRectangle(circRightShape, Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(1)), "CircleToSquareRight", NumSRs.GetNumRobustnessSRs.ToString)
                                    CreateOverview.Update_Overview()
                                    MsgBox("We have added new instances to Stage 5 Argument Patterns" + vbCrLf + "and updated the Multi-Overview to reflect the additions.")
                                End If
                            Catch ex As Exception
                                MsgBox("Please enter a number between 0 to 6.")
                            End Try

                        End If
                        NumSRs.WriteSRsToFile()
                    End If
                    'If the user clicked the diamond then process their responses
                    'We will remove the left or right branch of stage 5
                    If ((clickedWindow.Application.AlertResponse = 0) And isDiamond) Then
                        'Dim strMasterNames As String() = Nothing
                        'Globals.ThisAddIn.Application.ActiveDocument.Pages.GetNamesU(strMasterNames)
                        'Dim intLowerBound As Integer = LBound(strMasterNames)
                        'Dim intUpperBound As Integer = UBound(strMasterNames)
                        'Debug.Print(Globals.ThisAddIn.Application.ActiveDocument.Name + " Lower bound:" + intLowerBound.ToString + " Upper bound:" + intUpperBound.ToString + vbCrLf)

                        'While intLowerBound <= intUpperBound

                        '    Debug.Print(strMasterNames(intLowerBound) + vbCrLf)
                        '    intLowerBound = intLowerBound + 1

                        'End While

                        'Ask the user which branch (below the diamond) they want deleted
                        ' and delete it.
                        Dim getMessage, Title, DefaultVal, MyValue
                        getMessage = "Do you want to REMOVE the left or right branch?" + vbCrLf + "Input 'left' or 'right' or click Cancel to cancel the action."     ' Set prompt.
                        Title = "Stage 5: Branching"    ' Set title.
                        DefaultVal = ""    ' Set default.
                        ' Display message, title, and default value.
                        MyValue = InputBox(getMessage, Title, DefaultVal)
                        Debug.Write("You input " + MyValue + ".")
                        If ((Not MyValue Is Nothing) And (Not MyValue = "")) Then
                            If ((MyValue.ToString.Trim.ToLower = "left") Or (MyValue.ToString.Trim.ToLower = "right")) Then
                                'Get the shape (arrow) that points to the branch we want to remove
                                Dim shp As Visio.Shape = ThisAddIn.ListNextConnections(Globals.ThisAddIn.Application.ActivePage, targetShape, MyValue.ToString.Trim.ToLower)
                                If (Not shp Is Nothing) Then
                                    'Post-order depth-first traversal to remove the branch we want deleted
                                    'Dim dfs As New DFS()
                                    'Count maximum number of shapes we could traverse (how many shapes on this page?).
                                    Dim max As Integer = Globals.ThisAddIn.Application.ActivePage.Shapes.Count
                                    Debug.Write("Depth First Search: ")
                                    DFS.DepthFirstSearch(Globals.ThisAddIn.Application.ActivePage, max, Globals.ThisAddIn.Application.ActivePage.Shapes, shp)

                                    'Tidy up (reset) counters
                                    DFS.ResetVertexCount()
                                    DFS.ResetNodeCount()

                                    'Delete the dangling arrow - below the diamond
                                    Dim selToDel As Visio.Selection
                                    selToDel = Globals.ThisAddIn.Application.ActivePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty)
                                    Call selToDel.Select(shp, Visio.VisSelectArgs.visSelect)
                                    selToDel.Delete()

                                    'Swap filled diamond (choice) with empty diamond
                                    'Dim visioDocs As Visio.Documents = Globals.ThisAddIn.Application.Documents
                                    'Dim rep As Object = visioDocs("AMLAS Tool Stencil.vssx").Masters("Clear Diamond")
                                    'targetShape.ReplaceShape(rep)
                                    'Dim page As String = Globals.ThisAddIn.Application.ActivePage.NameU
                                    ReplaceShapeWithRectangle(targetShape, Globals.ThisAddIn.Application.ActivePage, "DiamondToSquare", Nothing)
                                    CreateOverview.Update_Overview()
                                    MsgBox("We have removed the required branch" + vbCrLf + "and updated the Multi-Overview to reflect the removals.")

                                End If
                            Else
                                MsgBox("Please enter either 'left' to remove the left branch or 'right' to remove the right branch.")
                            End If

                        End If
                        NumSRs.WriteSRsToFile()
                    End If
                    'If the user clicked the update overview button then process their responses
                    'We will update the multi-overview
                    If ((clickedWindow.Application.AlertResponse = 0) And isUpdateOverview) Then
                        Try
                            CreateOverview.Update_Overview()
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                        NumSRs.WriteSRsToFile()
                    End If
                Else
                    If (clickedWindow.Application.AlertResponse = 0) Then
                        'System.Windows.Forms.MessageBox.Show(
                        'noShapesFoundPrompt)
                    End If
                End If
            End If

        Catch errorThrown As System.Runtime.InteropServices.COMException
            Debug.WriteLine(errorThrown.Message)
        End Try

    End Sub
End Class
