Imports Microsoft.VisualBasic

Friend Class CreateOverview
    'Create the Overview pages from the 6 argument pages.

    'Get the shapes inside a bounding rectangle
    'For selecting subsets on pages
    Private Shared Function SelectByRectangularCrossingBox(page As String, lowerleftX As Double, lowerleftY As Double, upperrightX As Double, upperrightY As Double) As Visio.Selection

        Dim scopeId As Integer = Globals.ThisAddIn.Application.BeginUndoScope("try")

        Dim rc As Visio.Shape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(page).DrawRectangle(lowerleftX, lowerleftY, upperrightX, upperrightY)
        Dim selected As Visio.Selection = rc.SpatialNeighbors(Visio.VisSpatialRelationCodes.visSpatialContain, 0.01, Visio.VisSpatialRelationFlags.visSpatialIncludeContainerShapes + Visio.VisSpatialRelationFlags.visSpatialIncludeDataGraphics)

        Globals.ThisAddIn.Application.EndUndoScope(scopeId, False)
        Return selected
    End Function

    Private Shared Sub GetContainedShapes(page As String, shapes As Visio.Shapes, lowerleftX As Double, lowerleftY As Double, upperrightX As Double, upperrightY As Double, shapesList As List(Of Visio.Shape))

        'Dim scopeId As Integer = Globals.ThisAddIn.Application.BeginUndoScope("try")

        'Dim rc As Visio.Shape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(page).DrawRectangle(lowerleftX + 0.1, lowerleftY + 0.16, upperrightX - 0.16, upperrightY - 0.3)
        Dim rc As Visio.Shape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(page).DrawRectangle(lowerleftX, lowerleftY, upperrightX, upperrightY)
        'Dim selected As Visio.Selection
        'Dim shapesList As New List(Of Visio.Shape)

        'selected = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(page).CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty)
        For Each shp As Visio.Shape In Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(page).Shapes
            Dim intReturnValue As Integer = rc.SpatialRelation(shp, 0.01, Visio.VisSpatialRelationFlags.visSpatialIncludeContainerShapes + Visio.VisSpatialRelationFlags.visSpatialIncludeDataGraphics)
            '= rc.SpatialNeighbors(Visio.VisSpatialRelationCodes.visSpatialContain, 0.01, Visio.VisSpatialRelationFlags.visSpatialIncludeContainerShapes + Visio.VisSpatialRelationFlags.visSpatialIncludeDataGraphics)
            If intReturnValue = Visio.VisSpatialRelationCodes.visSpatialContain Then
                'Call selected.Select(shp, Visio.VisSelectArgs.visSelect)
                shapesList.Add(shp)
            End If
        Next
        'Globals.ThisAddIn.Application.EndUndoScope(scopeId, False)
        rc.Delete()
        'Return shapesList
    End Sub

    Private Shared Function ReplaceString(vsoCharacters1 As Visio.Characters, num As Integer, suffix As String) As String
        'Dim vsoCharacters1 As Visio.Characters
        Dim vsoStrng, vsoFind, vsoReplace As String
        'vsoCharacters1 = srcShape44.Characters
        vsoStrng = vsoCharacters1.Text
        vsoFind = "(5[.]\d+)[.]*\d*[R,P]*"
        vsoReplace = "$1." + num.ToString + suffix
        'result = Regex.Replace(Of String, "(\d+[A-Z]+) ", "$1|")
        Dim newString As String = System.Text.RegularExpressions.Regex.Replace(vsoStrng, vsoFind, vsoReplace)
        Return newString
    End Function
    'The code is wrapped around with BeginUndoScope/EndUndoScope to cancel changes.

    Private Shared Sub GetShapeLists(shapesList As List(Of Visio.Shapes), shpList As List(Of Visio.Shape), coordsList As List(Of System.Drawing.Point), offset As Double)
        Dim pwidth As Double
        Dim pheight As Double
        Dim bwidth As Double
        Dim bheight As Double
        Dim dbTop As Double
        Dim dbBottom As Double
        Dim dbLeft As Double
        Dim dbRight As Double

        Dim o1Shapes = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(0)).Shapes
        shapesList.Add(o1Shapes)
        Dim border As Visio.Shape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(0)).Shapes("Banner.93")
        border.BoundingBox(Visio.VisBoundingBoxArgs.visBBoxUprightWH, dbLeft, dbBottom, dbRight, dbTop)
        'Find the size of the selections bounding box.
        'Get the width and height of the bounding box
        pwidth = dbRight - dbLeft
        pheight = dbTop - dbBottom
        System.Diagnostics.Debug.Write("Stage 1: width, height " + pwidth.ToString + "," + pheight.ToString + vbCrLf)
        Dim p1 As New System.Drawing.Point(17.0 + (offset / 2), 17.0)
        coordsList.Add(p1)

        Dim o2Shapes = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(1)).Shapes
        shapesList.Add(o2Shapes)
        border = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(1)).Shapes("Banner.93")
        border.BoundingBox(Visio.VisBoundingBoxArgs.visBBoxUprightWH, dbLeft, dbBottom, dbRight, dbTop)
        'Find the size of the selections bounding box.
        'Get the width and height of the bounding box
        pwidth = dbRight - dbLeft
        pheight = dbTop - dbBottom
        System.Diagnostics.Debug.Write("Stage 2: width, height " + pwidth.ToString + "," + pheight.ToString + vbCrLf)
        Dim p2 As New System.Drawing.Point(8.8 + (offset / 2), 4.6)
        coordsList.Add(p2)

        Dim o3Shapes = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(2)).Shapes
        shapesList.Add(o3Shapes)
        border = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(2)).Shapes("Banner.93")
        border.BoundingBox(Visio.VisBoundingBoxArgs.visBBoxUprightWH, dbLeft, dbBottom, dbRight, dbTop)
        'Find the size of the selections bounding box.
        'Get the width and height of the bounding box
        pwidth = dbRight - dbLeft
        pheight = dbTop - dbBottom
        System.Diagnostics.Debug.Write("Stage 3: width, height " + pwidth.ToString + "," + pheight.ToString + vbCrLf)
        Dim p3 As New System.Drawing.Point(-0.9, -2.7)
        coordsList.Add(p3)

        Dim o4Shapes = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(3)).Shapes
        shapesList.Add(o4Shapes)
        border = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(3)).Shapes("Classic")
        border.BoundingBox(Visio.VisBoundingBoxArgs.visBBoxUprightWH, dbLeft, dbBottom, dbRight, dbTop)
        'Find the size of the selections bounding box.
        'Get the width and height of the bounding box
        pwidth = dbRight - dbLeft
        pheight = dbTop - dbBottom
        System.Diagnostics.Debug.Write("Stage 4: width, height " + pwidth.ToString + "," + pheight.ToString + vbCrLf)
        Dim p4 As New System.Drawing.Point(17.4 + offset, -3.3)
        coordsList.Add(p4)

        'http://visguy.com/vgforum/index.php?topic=7868.0
        'http://visguy.com/vgforum/index.php?topic=6595.0
        Dim o5Shapes = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(4)).Shapes
        shapesList.Add(o5Shapes)
        border = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(4)).Shapes("Classic")
        border.BoundingBox(Visio.VisBoundingBoxArgs.visBBoxUprightWH, dbLeft, dbBottom, dbRight, dbTop)
        'https://superuser.com/questions/1312687/visio-stop-shapes-from-changing-shapes-shape
        Dim v1 As Integer = border.CellsU("LockAspect").ResultIU
        Dim v2 As Integer = border.CellsU("LockHeight").ResultIU
        Dim v3 As Integer = border.CellsU("LockWidth").ResultIU
        Dim v4 As Integer = border.CellsU("LockFormat").ResultIU

        System.Diagnostics.Debug.Write("Overview before: " + dbLeft.ToString + ", " + dbBottom.ToString + ", " + dbRight.ToString + ", " + dbTop.ToString + ", lockformat " + border.CellsU("LockFormat").ResultIU.ToString + vbCrLf)
        'Dim selectedShapes As Visio.Selection = SelectByRectangularCrossingBox("Assurance Argument Pattern for ML Verification", dblLeft, dblBottom, dblRight, dblTop)
        bwidth = border.Cells("width").ResultIU
        bheight = border.Cells("height").ResultIU
        GetContainedShapes(ThisAddIn.stageNames(4), o5Shapes, dbLeft, dbBottom, dbRight, dbTop, shpList)

        border.CellsU("LockAspect").ResultIU = 0
        border.CellsU("LockHeight").ResultIU = 0
        border.CellsU("LockWidth").ResultIU = 0
        border.CellsU("LockFormat").ResultIU = 0
        border.Cells("width").ResultIU = bwidth - 0.019685
        border.Cells("height").ResultIU = bheight - 0.15748
        border.CellsU("LockAspect").ResultIU = v1
        border.CellsU("LockHeight").ResultIU = v2
        border.CellsU("LockWidth").ResultIU = v3
        border.CellsU("LockFormat").ResultIU = v4

        Dim rc2 As Visio.Shape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(4)).Shapes("Classic")
        rc2.BoundingBox(Visio.VisBoundingBoxArgs.visBBoxUprightWH, dbLeft, dbBottom, dbRight, dbTop)
        System.Diagnostics.Debug.Write("Overview2 after: " + dbLeft.ToString + ", " + dbBottom.ToString + ", " + dbRight.ToString + ", " + dbTop.ToString + ", lockformat " + rc2.CellsU("LockFormat").ResultIU.ToString + vbCrLf)

        border.BoundingBox(Visio.VisBoundingBoxArgs.visBBoxUprightWH, dbLeft, dbBottom, dbRight, dbTop)
        System.Diagnostics.Debug.Write("Overview1 after: " + dbLeft.ToString + ", " + dbBottom.ToString + ", " + dbRight.ToString + ", " + dbTop.ToString + ", lockformat " + border.CellsU("LockFormat").ResultIU.ToString + vbCrLf)

        shpList.Add(border)
        'Find the size of the selections bounding box.
        'Get the width and height of the bounding box
        pwidth = dbRight - dbLeft
        pheight = dbTop - dbBottom
        System.Diagnostics.Debug.Write("Stage 5: width, height " + pwidth.ToString + "," + pheight.ToString + vbCrLf)
        Dim p5 As New System.Drawing.Point(8.0, -6.1)
        coordsList.Add(p5)

        Dim o6Shapes = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(5)).Shapes
        shapesList.Add(o6Shapes)
        border = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(5)).Shapes("Classic")
        border.BoundingBox(Visio.VisBoundingBoxArgs.visBBoxUprightWH, dbLeft, dbBottom, dbRight, dbTop)
        'Find the size of the selections bounding box.
        'Get the width and height of the bounding box
        pwidth = dbRight - dbLeft
        pheight = dbTop - dbBottom
        System.Diagnostics.Debug.Write("Stage 6: width, height " + pwidth.ToString + "," + pheight.ToString + vbCrLf)
        Dim p6 As New System.Drawing.Point(27.7 + offset, 3.4)
        coordsList.Add(p6)

    End Sub
    Public Shared Sub Create_Overview(ByVal howMany As Integer, ByVal isLeftCircle As Boolean)
        Dim widthOffset As Double = 10.5
        Dim offset As Double
        If isLeftCircle Then
            offset = (howMany - 1) * widthOffset
        Else
            offset = (howMany - 1) * widthOffset + (NumSRs.GetNumPerformanceSRs() * widthOffset)
        End If

        Dim suffix As String
        If isLeftCircle Then
            suffix = "P"
        Else
            suffix = "R"
        End If

        Dim stageMultiOverview As Integer = 6
        Call ThisAddIn.DeleteShapes_Selection(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)))

        'Process each stage, get all shapes from the stage in a list, get the height and width of the stage shapes bounding box for layout
        Dim strMasterNames() As String = Nothing 'dummy so we can see the list of pages
        Globals.ThisAddIn.Application.ActiveDocument.Pages.GetNamesU(strMasterNames)

        Dim shapesList As New List(Of Visio.Shapes)
        Dim coordsList As New List(Of System.Drawing.Point)
        Dim shpList As New List(Of Visio.Shape)
        GetShapeLists(shapesList, shpList, coordsList, offset)

        ' Check if any shapes were found.
        Dim index As Int16 = 0
        For Each oShapes As Visio.Shapes In shapesList
            Dim point = coordsList.ElementAt(index)
            If ((Not oShapes Is Nothing) And
                                    (oShapes.Count > 0)) Then

                'message.Append(shapesFoundPrompt & System.Environment.NewLine)

                ' Loop through the collection of clicked shapes.
                ' Section 5 (index = 4) may be pasted multiple times
                'https://docs.microsoft.com/en-us/office/vba/api/visio.shape.group
                'https://stackoverflow.com/questions/32405404/vba-change-the-color-of-a-rounded-rectangle-in-visio
                'https://www.office-forums.com/threads/re-shape-protection-and-or-layer-properties-prevent-complete-execution-of-this-command.49668/
                If (4 = index) Then 'Stage 5 - treat this differently
                    Dim outline As Visio.Shape = Nothing
                    Dim keep As New List(Of Visio.Shape)
                    'If we are doing robustness SRs ands we already have performance SRs
                    ' then we need to add the robustness SRs back in first
                    If ((Not isLeftCircle) And (NumSRs.GetNumPerformanceSRs() > 0)) Then
                        'Copy and paste shape to the Overview.
                        'We exclude some shapes we do not want to copy.
                        For nextShape = 1 To oShapes.Count
                            If (Not oShapes.Item(nextShape).Name.Trim().StartsWith("Banner") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("Classic") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("Buttons") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("References")) Then
                                Dim srcShape4, dstShape4 As Visio.Shape
                                srcShape4 = (oShapes.Item(nextShape))
                                Dim cellX As Double = srcShape4.Cells("PinX").ResultIU
                                Dim cellY As Double = srcShape4.Cells("PinY").ResultIU
                                Dim move As Double
                                move = 0

                                'Paste into Overview

                                dstShape4 = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)).Drop(srcShape4, point.X + cellX + move, point.Y + cellY)

                                'Connect the top shape from this stage to the circle where we chose how many copies we wanted
                                If dstShape4.Name.Trim().StartsWith("Instantiable_goal_5_1") Then
                                    Dim cShape As Visio.Shape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)).Shapes("Dynamic connector")
                                    Dim fShape As Visio.Shape = Nothing
                                    'If isLeftCircle Then
                                    fShape = ThisAddIn.GetShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)), "CircleLeft")
                                    If (fShape Is Nothing) Then
                                        fShape = ThisAddIn.GetShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)), "CircleToSquareLeft")
                                    End If
                                    fShape.AutoConnect(dstShape4, Visio.VisAutoConnectDir.visAutoConnectDirNone, cShape)
                                    'cShape.CellsU("LinePattern").ResultIU = 23
                                End If

                            End If
                        Next nextShape
                    End If
                    ' Add the shpaes to the overview page in the correct place - with offsetting
                    ' if we have multiple SRs
                    For i As Integer = 1 To howMany
                        'Copy and paste shape to the Overview.
                        'We exclude some shapes we do not want to copy.
                        For nextShape = 0 To shpList.Count - 1
                            If (Not shpList.Item(nextShape).Name.Trim().StartsWith("Banner") And Not shpList.Item(nextShape).Name.Trim().StartsWith("Classic") And Not shpList.Item(nextShape).Name.Trim().StartsWith("Buttons") And Not shpList.Item(nextShape).Name.Trim().StartsWith("References")) Then
                                Dim srcShape4, dstShape4 As Visio.Shape
                                srcShape4 = (shpList.Item(nextShape))
                                If i = 1 Then
                                    keep.Add(srcShape4)
                                End If
                                'srcShape44 = (oShapes.Item(nextShape))
                                Dim cellX As Double = srcShape4.Cells("PinX").ResultIU
                                Dim cellY As Double = srcShape4.Cells("PinY").ResultIU
                                Dim move As Double
                                If isLeftCircle Then
                                    move = (i - 1) * widthOffset
                                Else
                                    move = (i - 1) * widthOffset + (NumSRs.GetNumPerformanceSRs() * widthOffset)
                                End If

                                'Paste into Overview
                                'Dim stage6 As Integer = 6
                                dstShape4 = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)).Drop(srcShape4, point.X + cellX + move, point.Y + cellY)

                                'Amend the shape identifiers in the shape text to reflect how many copies we have
                                Dim currentVal As Integer = dstShape4.CellsU("LockTextEdit").ResultIU
                                dstShape4.CellsU("LockTextEdit").ResultIU = 0
                                Dim newString As String = ReplaceString(dstShape4.Characters, i, suffix)
                                dstShape4.Text = newString
                                dstShape4.CellsU("LockTextEdit").ResultIU = currentVal
                                'vsoCharacters1.Text = Replace(vsoStrng, vsoFind, vsoReplace)

                                'Connect the top shape from this stage to the circle where we chose how many copies we wanted
                                If dstShape4.Name.Trim().StartsWith("Instantiable_goal_5_1") Then
                                    Dim cShape As Visio.Shape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)).Shapes("Dynamic connector")
                                    Dim fShape As Visio.Shape
                                    If isLeftCircle Then
                                        fShape = ThisAddIn.GetShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)), "CircleLeft")
                                    Else
                                        fShape = ThisAddIn.GetShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)), "CircleRight")
                                    End If
                                    'Dim tShape As Visio.Shape = getShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU("Overview"), "Goal (Stage 2)")
                                    'circShape.AutoConnect(visioRectShape, VisAutoConnectDir.visAutoConnectDirNone, vsoConnectorShape)
                                    fShape.AutoConnect(dstShape4, Visio.VisAutoConnectDir.visAutoConnectDirNone, cShape)

                                End If
                                ''Connect the top shape from this stage to the circle where we chose how many copies we wanted
                                'If dstShape4.Name.Trim().StartsWith("Instantiable_goal_5_1") Then
                                '    Dim cShape As Visio.Shape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)).Shapes("Dynamic connector")
                                '    'cShape.NameU = "DynamicConnector2-5-" + i.ToString
                                '    Dim fShape As Visio.Shape
                                '    Dim fShapeTmp As Visio.Shape
                                '    If isLeftCircle Then
                                '        fShapeTmp = ThisAddIn.getShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)), "CircleLeft")
                                '        If fShapeTmp IsNot Nothing Then
                                '            HandleMouseEvents.ReplaceShapeWithRectangle(fShapeTmp, Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)), "CircleToSquareLeft", ThisAddIn.getNumPerformanceSRs().ToString)
                                '        End If
                                '        fShape = ThisAddIn.getShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)), "CircleToSquareLeft")

                                '    Else
                                '        fShapeTmp = ThisAddIn.getShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)), "CircleRight")
                                '        If fShapeTmp IsNot Nothing Then
                                '            HandleMouseEvents.ReplaceShapeWithRectangle(fShapeTmp, Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)), "CircleToSquareRight", ThisAddIn.getNumRobustnessSRs().ToString)
                                '        End If
                                '        fShape = ThisAddIn.getShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)), "CircleToSquareRight")
                                '    End If
                                '    'Dim tShape As Visio.Shape = getShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU("Overview"), "Goal (Stage 2)")
                                '    'circShape.AutoConnect(visioRectShape, VisAutoConnectDir.visAutoConnectDirNone, vsoConnectorShape)
                                '    fShape.AutoConnect(dstShape4, Visio.VisAutoConnectDir.visAutoConnectDirNone, cShape)
                                '    'cShape.CellsU("LinePattern").ResultIU = 23
                                'End If
                                'Paste into stage 5 Assurance Argument Pattern
                                'We alread have it once so paste (howMany - 1) times
                                'If (i > 1) Then
                                'dstShape44 = 
                                'Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU("Assurance Argument Pattern for ML Verification").Drop(srcShape44, cellX + move, cellY)
                                'ThisAddIn.Application_ShapeFill(dstShape44, 5)
                                'End If

                            End If
                            'get a copy of the outlining rectangle and text
                            If (shpList.Item(nextShape).Name.Trim().StartsWith("Classic")) Then
                                Dim s As String = oShapes.Item(nextShape).Name.Trim()
                                outline = shpList.Item(nextShape)
                            End If
                        Next nextShape
                    Next


                    ' Paste additional copies of stage 5 to reflect how many performance and robustness copies we need
                    ' Relabel text in shapes to reflect number of copies of stage 5
                    ' These are copy instances
                    Dim todo As Integer
                    If isLeftCircle Then
                        todo = 2
                    Else
                        If NumSRs.GetNumPerformanceSRs = 0 Then
                            todo = 2
                        Else
                            todo = 1
                        End If

                    End If
                    For i As Integer = todo To howMany
                        'Copy and paste shape to the Overview.
                        'We exclude some shapes we do not want to copy.
                        For nextShape = 0 To keep.Count - 1
                            If (Not keep.Item(nextShape).Name.Trim().StartsWith("Banner") And Not keep.Item(nextShape).Name.Trim().StartsWith("Classic") And Not keep.Item(nextShape).Name.Trim().StartsWith("Buttons") And Not keep.Item(nextShape).Name.Trim().StartsWith("References")) Then
                                Dim srcShape44, dstShape44 As Visio.Shape
                                srcShape44 = (keep.Item(nextShape))
                                Dim cellX As Double = srcShape44.Cells("PinX").ResultIU
                                Dim cellY As Double = srcShape44.Cells("PinY").ResultIU
                                Dim move As Double
                                If isLeftCircle Then
                                    move = (i - 1) * widthOffset
                                Else
                                    move = (i - 1) * widthOffset + (NumSRs.GetNumPerformanceSRs() * widthOffset)
                                End If

                                'Paste into stage 5 Assurance Argument Pattern
                                'We alread have it once so paste (howMany - 1) times
                                'If (i > 1) Then
                                dstShape44 = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(index)).Drop(srcShape44, cellX + move, cellY)
                                Dim currentVal As Integer = dstShape44.CellsU("LockTextEdit").ResultIU
                                dstShape44.CellsU("LockTextEdit").ResultIU = 0
                                Dim newString As String = ReplaceString(dstShape44.Characters, i, suffix)
                                dstShape44.Text = newString
                                dstShape44.CellsU("LockTextEdit").ResultIU = currentVal
                                'vsoCharacters1.Text = Replace(vsoStrng, vsoFind, vsoReplace)

                                'ThisAddIn.Application_ShapeFill(dstShape44, 5)
                                'End If

                            End If

                        Next nextShape
                        ' Paste a copy of the outlining rectangle and text
                        If (Not outline Is Nothing) Then
                            Dim cellX As Double = outline.Cells("PinX").ResultIU
                            Dim cellY As Double = outline.Cells("PinY").ResultIU
                            Dim move As Double
                            If isLeftCircle Then
                                move = (i - 1) * widthOffset
                            Else
                                move = (i - 1) * widthOffset + (NumSRs.GetNumPerformanceSRs() * widthOffset)
                            End If
                            Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(index)).Drop(outline, cellX + move, cellY)
                        End If

                    Next
                    'If this is the leftmost box of shapes - we are processing the left circle
                    ' or we are processing the right circle with no performance SRs
                    If isLeftCircle Or ((Not isLeftCircle) And (NumSRs.GetNumPerformanceSRs() = 0)) Then
                        ' Relabel text in shapes to reflect number of copies of stage 5
                        ' This is the original instance - not a copy
                        Dim visSelection As Visio.Selection
                        Globals.ThisAddIn.Application.ActiveWindow.DeselectAll()
                        visSelection = Globals.ThisAddIn.Application.ActiveWindow.Selection
                        For nextShape = 0 To keep.Count - 1
                            If (Not keep.Item(nextShape).Name.Trim().StartsWith("Banner") And Not keep.Item(nextShape).Name.Trim().StartsWith("Classic") And Not keep.Item(nextShape).Name.Trim().StartsWith("Buttons") And Not keep.Item(nextShape).Name.Trim().StartsWith("References")) Then
                                Dim srcShape44 As Visio.Shape
                                srcShape44 = (keep.Item(nextShape))
                                Dim currentVal As Integer = srcShape44.CellsU("LockTextEdit").ResultIU
                                srcShape44.CellsU("LockTextEdit").ResultIU = 0
                                Dim newString As String = ReplaceString(srcShape44.Characters, 1, suffix)
                                srcShape44.Text = newString
                                srcShape44.CellsU("LockTextEdit").ResultIU = currentVal
                                'vsoCharacters1.Text = Replace(vsoStrng, vsoFind, vsoReplace)

                                'ThisAddIn.Application_ShapeFill(dstShape44, 5)
                                'End If
                                visSelection.Select(srcShape44, Visio.VisSelectArgs.visSelect)

                            End If

                        Next nextShape
                    End If

                Else
                    'All stages except stage 5
                    'Dim vsoSelection As Visio.Selection
                    'vsoSelection = Globals.ThisAddIn.Application.ActiveWindow.Selection
                    'Select all the objects on the page.
                    'https://microsoft.public.visio.developer.vba.narkive.com/8Rfju4rU/size-to-fit-drawing-contents
                    'Get the current selection.
                    Dim visSelection As Visio.Selection
                    Globals.ThisAddIn.Application.ActiveWindow.DeselectAll()
                    visSelection = Globals.ThisAddIn.Application.ActiveWindow.Selection
                    'Dim vsoShape As Visio.Shape
                    'https://www.dotnetperls.com/keyvaluepair-vbnet
                    'Dim adjMatrix As List(Of KeyValuePair(Of Visio.Shape, Visio.Shape)) = New List(Of KeyValuePair(Of Visio.Shape, Visio.Shape))
                    For nextShape = 1 To oShapes.Count
                        'Copy and paste shape to the Overview.
                        'We exclude some shapes we do not want to copy.
                        If (Not oShapes.Item(nextShape).Name.Trim().StartsWith("Dangling") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("Banner") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("Classic") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("Buttons") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("References") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("Dangling")) Then
                            Dim srcShape, dstShape As Visio.Shape
                            srcShape = (oShapes.Item(nextShape))

                            visSelection.Select(srcShape, Visio.VisSelectArgs.visSelect)
                            Dim cellX As Double = srcShape.Cells("PinX").ResultIU
                            Dim cellY As Double = srcShape.Cells("PinY").ResultIU
                            dstShape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(6)).Drop(srcShape, point.X + cellX, point.Y + cellY)

                            'https://stackoverflow.com/questions/56809118/update-the-fill-colour-of-each-shape-immediately-after-it-is-changed-in-a-loop
                            'https://stackoverflow.com/questions/37413207/visio-change-color-of-all-child-elements-using-vba

                            'Dim lngShapeIDs() As Integer
                            'lngShapeIDs = srcShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "")
                            'For intCount = 0 To UBound(lngShapeIDs)
                            '    adjMatrix.Add(New KeyValuePair(Of Visio.Shape, Visio.Shape)(srcShape, Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(index)).Shapes.ItemFromID(lngShapeIDs(intCount))))
                            'Next

                        End If
                    Next nextShape
                    'https://stackoverflow.com/questions/46972697/ms-visio-trying-to-check-the-end-arrow-type-for-a-connection-between-2-shapes
                End If
            End If
            index += 1
        Next

        TidyShapes()
        AddButton()

    End Sub
    Public Shared Sub Update_Overview()
        Dim howmany As Integer = NumSRs.GetNumPerformanceSRs() + NumSRs.GetNumRobustnessSRs()
        Dim widthOffset As Double = 10.5
        Dim offset As Double = 0.0
        offset = (NumSRs.GetNumRobustnessSRs() - 1) * widthOffset + (NumSRs.GetNumPerformanceSRs() * widthOffset)
        Dim suffix As String

        Dim stageMultiOverview As Integer = 6
        Call ThisAddIn.DeleteShapes_Selection(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)))

        'Process each stage, get all shapes from the stage in a list, get the height and width of the stage shapes bounding box for layout
        Dim strMasterNames() As String = Nothing 'dummy so we can see the list of pages
        Globals.ThisAddIn.Application.ActiveDocument.Pages.GetNamesU(strMasterNames)

        Dim shapesList As New List(Of Visio.Shapes)
        Dim coordsList As New List(Of System.Drawing.Point)
        Dim shpList As New List(Of Visio.Shape)
        GetShapeLists(shapesList, shpList, coordsList, offset)

        ' Check if any shapes were found.
        Dim index As Int16 = 0
        For Each oShapes As Visio.Shapes In shapesList
            Dim point = coordsList.ElementAt(index)
            If ((Not oShapes Is Nothing) And
                                    (oShapes.Count > 0)) Then

                'message.Append(shapesFoundPrompt & System.Environment.NewLine)

                ' Loop through the collection of clicked shapes.
                ' Section 5 (index = 4) may be pasted multiple times
                'https://docs.microsoft.com/en-us/office/vba/api/visio.shape.group
                'https://stackoverflow.com/questions/32405404/vba-change-the-color-of-a-rounded-rectangle-in-visio
                'https://www.office-forums.com/threads/re-shape-protection-and-or-layer-properties-prevent-complete-execution-of-this-command.49668/
                If (4 = index) Then 'Stage 5 - treat this differently
                    'Get the current selection.
                    Dim visSelection As Visio.Selection
                    Globals.ThisAddIn.Application.ActiveWindow.DeselectAll()
                    visSelection = Globals.ThisAddIn.Application.ActiveWindow.Selection
                    Dim goal_5_1_counter = 0
                    For nextShape = 1 To oShapes.Count
                        'Copy and paste shape to the Overview.
                        'We exclude some shapes we do not want to copy.
                        If (Not oShapes.Item(nextShape).Name.Trim().StartsWith("Dangling") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("Banner") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("Classic") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("Buttons") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("References") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("Dangling")) Then
                            Dim srcShape, dstShape As Visio.Shape
                            srcShape = (oShapes.Item(nextShape))

                            visSelection.Select(srcShape, Visio.VisSelectArgs.visSelect)
                            Dim cellX As Double = srcShape.Cells("PinX").ResultIU
                            Dim cellY As Double = srcShape.Cells("PinY").ResultIU
                            dstShape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)).Drop(srcShape, point.X + cellX, point.Y + cellY)

                            'https://stackoverflow.com/questions/56809118/update-the-fill-colour-of-each-shape-immediately-after-it-is-changed-in-a-loop
                            'https://stackoverflow.com/questions/37413207/visio-change-color-of-all-child-elements-using-vba

                            'Dim lngShapeIDs() As Integer
                            'lngShapeIDs = srcShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "")
                            'For intCount = 0 To UBound(lngShapeIDs)
                            '    adjMatrix.Add(New KeyValuePair(Of Visio.Shape, Visio.Shape)(srcShape, Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(index)).Shapes.ItemFromID(lngShapeIDs(intCount))))
                            'Next
                            'Connect the top shape from this stage to the circle where we chose how many copies we wanted
                            If dstShape.Name.Trim().StartsWith("Instantiable_goal_5_1") Then
                                Dim cShape As Visio.Shape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)).Shapes("Dynamic connector")
                                Dim fShape As Visio.Shape = Nothing
                                If goal_5_1_counter < NumSRs.GetNumPerformanceSRs Then
                                    fShape = ThisAddIn.GetShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)), "CircleToSquareLeft")
                                    If (fShape Is Nothing) Then
                                        fShape = ThisAddIn.GetShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)), "CircleLeft")
                                    End If
                                Else
                                    fShape = ThisAddIn.GetShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)), "CircleToSquareRight")
                                    If (fShape Is Nothing) Then
                                        fShape = ThisAddIn.GetShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)), "CircleRight")
                                    End If
                                End If
                                'Dim tShape As Visio.Shape = getShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU("Overview"), "Goal (Stage 2)")
                                'circShape.AutoConnect(visioRectShape, VisAutoConnectDir.visAutoConnectDirNone, vsoConnectorShape)
                                fShape.AutoConnect(dstShape, Visio.VisAutoConnectDir.visAutoConnectDirNone, cShape)
                                'cShape.CellsU("LinePattern").ResultIU = 23
                                goal_5_1_counter += 1

                            End If

                        End If
                    Next nextShape


                Else
                    'All stages except stage 5
                    'Dim vsoSelection As Visio.Selection
                    'vsoSelection = Globals.ThisAddIn.Application.ActiveWindow.Selection
                    'Select all the objects on the page.
                    'https://microsoft.public.visio.developer.vba.narkive.com/8Rfju4rU/size-to-fit-drawing-contents
                    'Get the current selection.
                    Dim visSelection As Visio.Selection
                    Globals.ThisAddIn.Application.ActiveWindow.DeselectAll()
                    visSelection = Globals.ThisAddIn.Application.ActiveWindow.Selection
                    'Dim vsoShape As Visio.Shape
                    'https://www.dotnetperls.com/keyvaluepair-vbnet
                    'Dim adjMatrix As List(Of KeyValuePair(Of Visio.Shape, Visio.Shape)) = New List(Of KeyValuePair(Of Visio.Shape, Visio.Shape))
                    For nextShape = 1 To oShapes.Count
                        'Copy and paste shape to the Overview.
                        'We exclude some shapes we do not want to copy.
                        If (Not oShapes.Item(nextShape).Name.Trim().StartsWith("Dangling") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("Banner") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("Classic") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("Buttons") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("References") And Not oShapes.Item(nextShape).Name.Trim().StartsWith("Dangling")) Then
                            Dim srcShape, dstShape As Visio.Shape
                            srcShape = (oShapes.Item(nextShape))

                            visSelection.Select(srcShape, Visio.VisSelectArgs.visSelect)
                            Dim cellX As Double = srcShape.Cells("PinX").ResultIU
                            Dim cellY As Double = srcShape.Cells("PinY").ResultIU
                            dstShape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)).Drop(srcShape, point.X + cellX, point.Y + cellY)

                            'https://stackoverflow.com/questions/56809118/update-the-fill-colour-of-each-shape-immediately-after-it-is-changed-in-a-loop
                            'https://stackoverflow.com/questions/37413207/visio-change-color-of-all-child-elements-using-vba

                            'Dim lngShapeIDs() As Integer
                            'lngShapeIDs = srcShape.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "")
                            'For intCount = 0 To UBound(lngShapeIDs)
                            '    adjMatrix.Add(New KeyValuePair(Of Visio.Shape, Visio.Shape)(srcShape, Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(index)).Shapes.ItemFromID(lngShapeIDs(intCount))))
                            'Next

                        End If
                    Next nextShape
                    'https://stackoverflow.com/questions/46972697/ms-visio-trying-to-check-the-end-arrow-type-for-a-connection-between-2-shapes
                End If
            End If
            index += 1
        Next
        TidyShapes()
        AddButton()
    End Sub

    Shared Sub TryAddImage(vPg As Integer, imageFileStr As String, x As Double, y As Double, width As Double, height As Double)
        Dim shpNew As Visio.Shape
        'shpNew = AddImageShape(vPg, imageFileStr)
        shpNew = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(vPg)).Import(imageFileStr)
        If Not shpNew Is Nothing Then
            'Do something to your new shape
            '//Set position
            shpNew.CellsU("PinX").ResultIU = x
            shpNew.CellsU("PinY").ResultIU = y
            '//Set size
            shpNew.CellsU("Width").ResultIU = width
            shpNew.CellsU("Height").ResultIU = height

        End If
    End Sub


    Private Shared Function AddImageShape(vPg As Integer, fileName As String) As Visio.Shape
        Dim vPage As Visio.Page = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(vPg))
        Dim shpNew As Visio.Shape = Nothing
        If Not vPage Is Nothing Then
            Dim UndoScopeID1 As Long
            UndoScopeID1 = Globals.ThisAddIn.Application.BeginUndoScope("Insert image shape")

            On Error Resume Next
            shpNew = vPage.Import(fileName)

            If Not shpNew Is Nothing Then
                Globals.ThisAddIn.Application.EndUndoScope(UndoScopeID1, True)
            Else
                Globals.ThisAddIn.Application.EndUndoScope(UndoScopeID1, False)
            End If
        End If

        AddImageShape = shpNew
        'Return shpNew
    End Function
    Private Shared Sub AddButton()
        Dim stageMultiOverview As Integer = 6
        Dim left As Double = 1.5
        Dim right As Double = 2.5
        Dim bottom As Double = 20
        Dim top As Double = 21
        Dim x As Double = (left + right) / 2
        Dim y As Double = (bottom + top) / 2
        Dim width As Double = right - left
        Dim height = top - bottom
        Dim rc As Visio.Shape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageMultiOverview)).DrawRectangle(left, bottom, right, top)
        rc.NameU = "Refresh"
        rc.CellsU("Fillforegnd").FormulaForceU = "RGB(211,211,211)"
        'This throws an error
        'Dim rStr As String = "Refresh and Update the Overview"
        'rc.CellsU("comment").FormulaForce = rStr

        'rc.ChangePicture("refresh_update.png")
        Dim fname As String = My.Computer.FileSystem.CurrentDirectory + "\\refresh_update.png"
        TryAddImage(stageMultiOverview, fname, x, y, width, height)
    End Sub
    Private Shared Sub TidyShapes()

        'Connect all shapes on stage 5 Assurance Argument Pattern for ML Verification page
        Reconnect(4)
        'Connect all shapes on Multi-Overview page
        Dim stageNameIndex As Integer = 6
        Reconnect(stageNameIndex)

        Dim connectShape As Visio.Shape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageNameIndex)).Shapes("Dynamic connector")
        Dim fromShape As Visio.Shape = ThisAddIn.GetShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageNameIndex)), "Strategy (Stage 1_1)")
        Dim toShape As Visio.Shape = ThisAddIn.GetShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageNameIndex)), "Goal (Stage 2_1)")
        'circShape.AutoConnect(visioRectShape, VisAutoConnectDir.visAutoConnectDirNone, vsoConnectorShape)
        fromShape.AutoConnect(toShape, Visio.VisAutoConnectDir.visAutoConnectDirNone, connectShape)

        Dim connectShape2 As Visio.Shape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageNameIndex)).Shapes("Dynamic connector")
        Dim toShape2 As Visio.Shape = ThisAddIn.GetShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageNameIndex)), "Goal stage 6_1")
        'circShape.AutoConnect(visioRectShape, VisAutoConnectDir.visAutoConnectDirNone, vsoConnectorShape)
        fromShape.AutoConnect(toShape2, Visio.VisAutoConnectDir.visAutoConnectDirNone, connectShape2)

        Dim connectShape3 As Visio.Shape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageNameIndex)).Shapes("Dynamic connector")
        Dim fromShape3 As Visio.Shape = ThisAddIn.GetShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageNameIndex)), "SquareLeft")
        Dim toShape3 As Visio.Shape = ThisAddIn.GetShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageNameIndex)), "Goal (Stage 3_1)")
        'circShape.AutoConnect(visioRectShape, VisAutoConnectDir.visAutoConnectDirNone, vsoConnectorShape)
        fromShape3.AutoConnect(toShape3, Visio.VisAutoConnectDir.visAutoConnectDirNone, connectShape3)

        Dim connectShape4 As Visio.Shape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageNameIndex)).Shapes("Dynamic connector")
        Dim fromShape4 As Visio.Shape = ThisAddIn.GetShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageNameIndex)), "SquareRight")
        Dim toShape4 As Visio.Shape = ThisAddIn.GetShapeByName(Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageNameIndex)), "Goal stage 4_1")
        'circShape.AutoConnect(visioRectShape, VisAutoConnectDir.visAutoConnectDirNone, vsoConnectorShape)
        fromShape4.AutoConnect(toShape4, Visio.VisAutoConnectDir.visAutoConnectDirNone, connectShape4)
    End Sub

    Private Shared Function GetAllSubShapes(ShpObj As Visio.Shape, SubShapes As List(Of Visio.Shape), Optional AddFirstShp As Boolean = False)
        If AddFirstShp Then SubShapes.Add(ShpObj)
        Dim CheckShp As Visio.Shape
        For Each CheckShp In ShpObj.Shapes
            SubShapes.Add(CheckShp)
            Call GetAllSubShapes(CheckShp, SubShapes, False)
        Next CheckShp
    End Function

    Private Shared Sub Reconnect(stageNameIndex As Integer)
        'http://visguy.com/vgforum/index.php?topic=5773.15
        Dim cons As Visio.Selection
        Dim con As Visio.Shape
        Dim shp As Visio.Shape
        Dim sel As Visio.Selection
        Dim xBeg As Double, xEnd As Double, yBeg As Double, yEnd As Double
        Dim xS As Double, yS As Double, xP As Double, yP As Double
        Dim i As Integer

        Const dXY = 0.1
        Const tolerance = 0.1

        Dim ocShapes = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageNameIndex)).Shapes
        cons = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU(ThisAddIn.stageNames(stageNameIndex)).CreateSelection(Visio.VisSelectionTypes.visSelTypeByRole, Visio.VisSelectMode.visSelModeSkipSuper, Visio.VisRoleSelectionTypes.visRoleSelConnector)

        For Each con In ocShapes
            If (con.OneD And (con.Connects.Count < 2)) Then
                xBeg = con.CellsU("BeginX").ResultIU
                yBeg = con.CellsU("BeginY").ResultIU
                xEnd = con.CellsU("EndX").ResultIU
                yEnd = con.CellsU("EndY").ResultIU


                sel = con.SpatialNeighbors(Visio.VisSpatialRelationCodes.visSpatialTouching, tolerance, 0)

                For Each shp In sel
                    For i = 0 To shp.RowCount(Visio.VisSectionIndices.visSectionConnectionPts) - 1
                        xS = shp.CellsSRC(Visio.VisSectionIndices.visSectionConnectionPts, i, Visio.VisCellIndices.visCnnctX).ResultIU
                        yS = shp.CellsSRC(Visio.VisSectionIndices.visSectionConnectionPts, i, Visio.VisCellIndices.visCnnctY).ResultIU
                        shp.XYToPage(xS, yS, xP, yP)

                        If (Math.Abs(xP - xBeg) < dXY) And (Math.Abs(yP - yBeg) < dXY) Then
                            con.CellsU("BeginX").GlueTo(shp.CellsSRC(Visio.VisSectionIndices.visSectionConnectionPts, i, 0))
                        End If
                        If (Math.Abs(xP - xEnd) < dXY) And (Math.Abs(yP - yEnd) < dXY) Then
                            con.CellsU("EndX").GlueTo(shp.CellsSRC(Visio.VisSectionIndices.visSectionConnectionPts, i, 0))
                        End If
                    Next i
                Next shp
            End If
        Next con
    End Sub
End Class
