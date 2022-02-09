Imports Microsoft.Office.Interop.Visio
Imports System.Data
Imports System.Runtime.InteropServices
Imports System.Diagnostics


Public Class ThisAddIn
    'global consts for colours
    Shared ReadOnly stage_colour(7) As String
    Shared stages(7) As String

    Friend Shared numPerformanceSRs As Integer = 0
    Friend Shared numRobustnessSRs As Integer = 0



    'Friend Shared stage1NameU As String = "Stage 1: Argument Pattern for ML Safety Assurance Scoping"
    'Friend Shared stage2NameU As String = "Stage 2: Assurance Argument Pattern for ML Safety Requirements"
    'Friend Shared stage3NameU As String = "Stage3: Assurance Argument Pattern for ML Data "
    'Friend Shared stage4NameU As String = "Stage 4:  Assurance Argument Pattern for Model Learning"
    'Friend Shared stage5NameU As String = "Assurance Argument Pattern for ML Verification"
    'Friend Shared stage6NameU As String = "Stage 6: Assurance Argument Pattern for ML Model Deployment"
    'Friend Shared stageMultiOverviewNameU As String = "Multi-Overview"

    Friend Shared stageNames() As String = {"Stage 1: Argument Pattern for ML Safety Assurance Scoping", "Stage 2: Assurance Argument Pattern for ML Safety Requirements", "Stage3: Assurance Argument Pattern for ML Data ", "Stage 4:  Assurance Argument Pattern for Model Learning", "Assurance Argument Pattern for ML Verification", "Stage 6: Assurance Argument Pattern for ML Model Deployment", "Multi-Overview"}




    'Add mouse event handler to window
    Private clickedWindow As Microsoft.Office.Interop.Visio.Window


    Private Sub Document_DocumentSaved(ByVal doc As IVDocument)
        NumSRs.WriteSRsToFile()

    End Sub

    Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
        NumSRs.WriteSRsToFile()

    End Sub


    Private Sub ThisAddIn_Startup() Handles Me.Startup
        'myMouseListener = New MouseListener

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
        'myMouseListener = Nothing
    End Sub
    <ComVisible(True)>
    Public Interface IAddInUtilities
        Sub Application_AddShape()
    End Interface


    '//https://bvisual.net/2009/09/16/listing-connections-in-visio-2010/
    '//http://visguy.com/vgforum/index.php?topic=3443.0
    Public Shared Function ListNextConnections(ByRef activePage As Microsoft.Office.Interop.Visio.Page, ByRef shp As Visio.Shape, myValue As String) As Visio.Shape
        'Dim shp As Visio.Shape
        'Dim connectorShape As Visio.Shape
        Dim outShape As Visio.Shape
        Dim targetShape As Visio.Shape = Nothing
        'Dim aryTargetIDs() As Long
        Dim outIDs() As Integer
        'Dim targetID As Long
        'Dim sourceID As Long
        Dim i As Integer
        Dim x As Integer

        If (shp Is Nothing) Then
            MsgBox("Select a shape with connections.")
            Return Nothing
        Else
            If ((myValue = "left") Or (myValue = "right")) Then
                If myValue = "left" Then
                    x = 99999
                Else
                    x = 0
                End If
                'If shp.CellExists("Prop.NetworkName", Visio.VisExistsFlags.visExistsAnywhere) Then
                Debug.Print("Shape", shp.Name)
                outIDs = shp.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "")
                'outIDs = shp.ConnectedShapes(VisConnectedShapesFlags.visConnectedShapesOutgoingNodes, "")
                For i = 0 To UBound(outIDs)
                    outShape = activePage.Shapes.ItemFromID(outIDs(i))
                    If Not outShape Is Nothing Then
                        If (myValue = "left" And outShape.Cells("PinX").ResultIU < x) Then
                            targetShape = outShape
                            x = outShape.Cells("PinX").ResultIU
                        End If
                        If (myValue = "right" And outShape.Cells("PinX").ResultIU > x) Then
                            targetShape = outShape
                            x = outShape.Cells("PinX").ResultIU
                        End If
                    End If
                Next

                'End If

                'Dim i As Integer

                'MsgBox "GluedShapes"
            Else
                MsgBox("Please enter 'left' or 'right' to remove that branch.")
                Return Nothing
            End If

        End If
        Return targetShape

    End Function

    ' Find a specified shape on a specified Visio page and return it
    ' Uses Shape.NameU (shape's universal name) to match shapes
    Public Shared Function GetShape(
            ByVal otherPage As Microsoft.Office.Interop.Visio.Page,
            ByVal otherShape As Visio.Shape) As Visio.Shape

        Dim clickedShape As Visio.Shape = Nothing

        Try

            ' Use the SpatialSearch method of the page to get the list
            ' of shapes at the location.
            Dim oShapes = otherPage.Shapes
            ' Check if any shapes were found.
            If ((Not oShapes Is Nothing) And
                        (oShapes.Count > 0)) Then

                'message.Append(shapesFoundPrompt & System.Environment.NewLine)

                ' Loop through the collection of clicked shapes.
                For nextShape = 1 To oShapes.Count
                    If (StrComp(otherShape.NameU, oShapes.Item(nextShape).NameU) = 0) Then
                        clickedShape = oShapes.Item(nextShape)
                        Return clickedShape
                    End If
                Next nextShape

            End If
        Catch errorThrown As System.Runtime.InteropServices.COMException
            Debug.WriteLine(errorThrown.Message)
        End Try

        Return clickedShape

    End Function
    ' Find a shape (using shape's name as a string) on a specified Visio page and return it
    ' Uses Shape.NameU (shape's universal name) to match shapes
    Public Shared Function GetShapeByName(
            ByVal otherPage As Microsoft.Office.Interop.Visio.Page,
            ByVal otherShape As String) As Visio.Shape

        Dim clickedShape As Visio.Shape = Nothing

        Try

            ' Use the SpatialSearch method of the page to get the list
            ' of shapes at the location.
            Dim oShapes = otherPage.Shapes
            ' Check if any shapes were found.
            If ((Not oShapes Is Nothing) And
                        (oShapes.Count > 0)) Then

                'message.Append(shapesFoundPrompt & System.Environment.NewLine)

                ' Loop through the collection of clicked shapes.
                For nextShape = 1 To oShapes.Count
                    If (oShapes.Item(nextShape).Name.Trim().StartsWith(otherShape)) Then
                        clickedShape = oShapes.Item(nextShape)
                        Return clickedShape
                    End If
                Next nextShape

            End If
        Catch errorThrown As System.Runtime.InteropServices.COMException
            Debug.WriteLine(errorThrown.Message)
        End Try

        Return clickedShape


    End Function

    ' Find a shape (using shape's name as a string) on a specified Visio page and return it
    ' Uses Shape.NameU (shape's universal name) to match shapes
    Public Shared Function GetShapeByID(
            ByVal otherPage As Microsoft.Office.Interop.Visio.Page,
            ByVal otherShape As Integer) As Visio.Shape

        Dim clickedShape As Visio.Shape = Nothing

        Try

            ' Use the SpatialSearch method of the page to get the list
            ' of shapes at the location.
            Dim oShapes = otherPage.Shapes
            ' Check if any shapes were found.
            If ((Not oShapes Is Nothing) And
                        (oShapes.Count > 0)) Then

                'message.Append(shapesFoundPrompt & System.Environment.NewLine)

                ' Loop through the collection of clicked shapes.
                For nextShape = 1 To oShapes.Count
                    If (oShapes.Item(nextShape).ID = otherShape) Then
                        clickedShape = oShapes.Item(nextShape)
                        Return clickedShape
                    End If
                Next nextShape

            End If
        Catch errorThrown As System.Runtime.InteropServices.COMException
            Debug.WriteLine(errorThrown.Message)
        End Try

        Return clickedShape


    End Function

    'Delete all shapes on a page
    'Disables shapes' protection prior to deletion.
    Public Shared Sub DeleteShapes_Selection(ByRef visPg As Visio.Page)

        '// Create an EMPTY selection of shapes (we'll add shapes
        '// to it later):   
        Dim selToDel As Visio.Selection
        selToDel = visPg.CreateSelection(
   Visio.VisSelectionTypes.visSelTypeEmpty)

        '// We can use For...Each again, because we're not deleting
        '// shapes from *within* the loop.
        Dim shp As Visio.Shape
        For Each shp In visPg.Shapes
            shp.CellsU("LockDelete").ResultIU = 0
            'shp.Protection.LockAspect.Value = False
            shp.CellsU("LockBegin").ResultIU = 0 'shp.Protection.LockBegin.Value = True
            'shp.Protection.LockCalcWH.Value = False
            'shp.Protection.LockCrop.Value = False
            'shp.Protection.LockCustProp.Value = False
            'shp.Protection.LockDelete.Value = False
            shp.CellsU("LockEnd").ResultIU = 0 'shp.Protection.LockEnd.Value = True
            'shp.Protection.LockFormat.Value = False
            'shp.Protection.LockFromGroupFormat.Value = False
            'shp.Protection.LockGroup.Value = False
            'shp.Protection.LockHeight.Value = False
            'shp.Protection.LockMoveX.Value = False
            'shp.Protection.LockMoveY.Value = False
            'shp.Protection.LockRotate.Value = False
            'shp.Protection.LockSelect.Value = False
            'shp.Protection.LockTextEdit.Value = False
            shp.CellsU("LockThemeColors").ResultIU = 0 'shp.Protection.LockThemeColors.Value = True
            shp.CellsU("LockThemeEffects").ResultIU = 0 'shp.Protection.LockThemeEffects.Value = True
            'shp.Protection.LockVtxEdit.Value = False
            '// Add shape to selection if it is 'wide', but don't
            '// whack the shape here:
            'If (m_filter_isWiderThan(shp, 25.4)) Then

            '// Annoying Visio-syntax for (simply) selecting
            '// a shape:
            Call selToDel.Select(shp, Visio.VisSelectArgs.visSelect)

            'End If

        Next

        '// Delete the selection, if it is not empty:
        If (selToDel.Count > 0) Then
            '// This is where shapes get whacked:
            selToDel.Delete()
        End If

        '// Cleanup:
        'selToDel = Nothing
        'shp = Nothing

    End Sub

    'Detect when the user clicks on somehting - mouseDown event
    'Calls diiferent functionality depending where they clicked
    Private Sub Application_MouseDown(ByVal mouseButton As Integer,
            ByVal buttonState As Integer,
            ByVal clickLocationX As Double,
            ByVal clickLocationY As Double,
            ByRef cancelProcessing As Boolean) Handles Application.MouseDown

        HandleMouseEvents.HandleMouseDown(mouseButton, buttonState, clickLocationX, clickLocationY, cancelProcessing, clickedWindow)
    End Sub
    'https://microsoft.public.visio.developer.vba.narkive.com/INia7TTY/why-don-t-shapesheet-changes-update-the-shape
    'http://visguy.com/vgforum/index.php?topic=8011.0
    Public Shared Sub Application_ShapeFill(Shape As Shape, index As Integer)
        'on drop, add stage colour to shape and prompt user to enter shape data
        Dim vsoMaster As Visio.Master
        Dim visioDocs As Visio.Documents = Globals.ThisAddIn.Application.Documents
        Dim activePage As String

        activePage = Globals.ThisAddIn.Application.ActivePage.Name
        'Get the Master property of the shape. 
        vsoMaster = Shape.Master

        'Check whether the shape has a master. If not, 
        'the shape was created locally. 
        'create background colour for dropped shapes
        For Each shapewithin As Visio.Shape In Shape.Shapes
            'If Not shapewithin.IsDataGraphicCallout Then
            shapewithin.CellsU("Fillforegnd").FormulaForceU = stage_colour(index)
            'End If
        Next shapewithin
        Shape.CellsU("Fillforegnd").FormulaForceU = stage_colour(index)

    End Sub
    'Public Shared Sub Application_ShapeAdded(Shape As Shape) Handles Application.ShapeAdded
    '    'on drop, add stage colour to shape and prompt user to enter shape data
    '    Dim vsoMaster As Visio.Master
    '    Dim visioDocs As Visio.Documents = Globals.ThisAddIn.Application.Documents
    '    Dim activePage As String

    '    activePage = Globals.ThisAddIn.Application.ActivePage.Name
    '    'Get the Master property of the shape. 
    '    vsoMaster = Shape.Master

    '    'Check whether the shape has a master. If not, 
    '    'the shape was created locally. 
    '    If vsoMaster IsNot Nothing Then
    '        'create background colour for dropped shapes
    '        For i = 1 To stages.Length - 1
    '            If activePage.Contains(stages(i)) Then
    '                If Not vsoMaster.Name.Contains("Document") Then 'documents have no colour
    '                    If vsoMaster.Name.Contains("Justification") Or vsoMaster.Name.Contains("Assumption") Or vsoMaster.Name.Contains("solution") Then
    '                        For Each shapewithin As Visio.Shape In Shape.Shapes
    '                            If Not shapewithin.IsDataGraphicCallout Then
    '                                'VJH - temp remove shapewithin.CellsU("Fillforegnd").FormulaU = stage_colour(i)
    '                            End If
    '                        Next shapewithin
    '                    Else
    '                        'VJH - temp remove Shape.CellsU("Fillforegnd").FormulaU = stage_colour(i)
    '                    End If
    '                End If
    '            End If

    '        Next i

    '    End If

    'End Sub

    Private Sub Application_ShapeChanged(Shape As Shape) Handles Application.ShapeChanged
        'Dim message As String
        Dim currentpage As String
        Dim correspondingArgPage As String
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
            'Shape.CellsU("LinePattern").FormulaU = "23"
            'Shape.CellsU("EndArrow").FormulaU = "13"
        End If

    End Sub

    Private Sub Application_FormulaChanged(Cell As Cell) Handles Application.FormulaChanged
        Dim d As String
        'formatting of desciption property on documents as requested by Richard Hawkins
        'if string entered for file path, strip {} but leave text

        'try to only strip if shape data field is changed
        If Cell.Name = "Prop.Inst_statement" And Cell.Shape.Name.Contains("Document") Then
            If Cell.FormulaU <> "" Then
                d = Cell.Shape.CellsU("Prop.Description").FormulaU
                d = d.Replace("}", "")
                d = d.Replace("{", "")
                Cell.Shape.CellsU("Prop.Description").FormulaU = d
            End If


        End If

    End Sub



    Private Sub Application_ViewChanged(Window As Window) Handles Application.ViewChanged
        'Globals.Ribbons.Ribbon1.prevPage = Globals.ThisAddIn.Application.ActivePage.Name
    End Sub

    Private Sub Application_BeforeModal(app As Application) Handles Application.BeforeModal
        'catch prevPage for button2 on ribbon after loading dialogue box to switch to page
        If Not IsNothing(Globals.ThisAddIn.Application.ActivePage) Then
            Globals.Ribbons.Ribbon1.prevPage = Globals.ThisAddIn.Application.ActivePage.Name
        End If
    End Sub


End Class


