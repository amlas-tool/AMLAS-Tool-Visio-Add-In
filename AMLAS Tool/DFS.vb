Imports Microsoft.VisualBasic
Imports System.Runtime.InteropServices
Imports System.Diagnostics

'https://www.programmingalgorithms.com/algorithm/depth-first-traversal/vb-net/
'Post-order depth-first search class for removing all shapes below a specified
'shape in the page's shape tree.
Friend Class DFS
	Public Structure Vertex
		Public shape As Visio.Shape
		Public Label As String
		Public Visited As Boolean
	End Structure

	'Private Shared _top As Integer = -1
	Private Shared _vertexCount As Integer = 0
	Private Shared _nodeCount As Integer = 0

	Public Shared Sub ResetVertexCount()
		_vertexCount = 0
	End Sub

	Public Shared Sub ResetNodeCount()
		_vertexCount = 0
	End Sub

	Private Shared Function PushObject(stack As Object(), item As Object, _top As Integer)
		_top += 1
		stack(_top) = item
		Return _top
	End Function

	Private Shared Function PopObject(stack As Object(), _top As Integer) As Object
		Dim retVal = stack(_top)
		Return retVal
	End Function

	Private Shared Function Push(stack As Integer(), item As Integer, _top As Integer)
		_top += 1
		stack(_top) = item
		Return _top
	End Function

	Private Shared Function Pop(stack As Integer(), _top As Integer) As Integer
		Dim retVal = stack(_top)
		Return retVal
	End Function

	Private Shared Function Peek(stack As Integer(), _top As Integer) As Integer
		Return stack(_top)
	End Function

	Private Shared Function IsStackEmpty(_top As Integer) As Boolean
		Return _top = -1
	End Function

	Public Shared Function AddVertex(arrVertices As Vertex(), label As String, _counter As Integer)
		Debug.Write("Adding: " + label + vbCrLf)
		Dim vertex As New Vertex()
		vertex.Label = label
		vertex.Visited = False
		arrVertices(_counter) = vertex
		_counter += 1
		Return _counter
	End Function


	Public Shared Function AddVertex(arrVertices As Vertex(), shape As Visio.Shape, _counter As Integer)
		Debug.Write("Adding: " + shape.NameU + vbCrLf)
		Dim vertex As New Vertex()
		vertex.shape = shape
		vertex.Label = shape.NameU
		vertex.Visited = False
		arrVertices(_counter) = vertex
		_counter += 1
		Return _counter
	End Function

	Public Shared Sub AddEdge(adjacencyMatrix As Integer(,), start As Integer, [end] As Integer)
		adjacencyMatrix(start, [end]) = 1
		adjacencyMatrix([end], start) = 1
	End Sub

	Private Shared Sub DisplayVertex(arrVertices As Vertex(), vertexIndex As Integer)
		Debug.Write(arrVertices(vertexIndex).Label & " ")
	End Sub

	Private Shared Function GetAdjacentUnvisitedVertex(arrVertices As Vertex(), adjacencyMatrix As Integer(,), vertexIndex As Integer) As Integer
		For i As Integer = 0 To _vertexCount - 1
			If adjacencyMatrix(vertexIndex, i) = 1 AndAlso arrVertices(i).Visited = False Then
				Return i
			End If
		Next
		Return -1
	End Function

	Private Shared Function GetAdjacentLeftVertex(arrVertices As Vertex(), adjacencyMatrix As Integer(,), vertexIndex As Integer) As Integer
		For i As Integer = 0 To _vertexCount - 1
			If adjacencyMatrix(vertexIndex, i) = 1 AndAlso arrVertices(i).Visited = False Then
				Return i
			End If
		Next
		Return -1
	End Function

	'Tidy the left hand branch of the tree (flowchart) on stage 5 argument pattern.
	Public Shared Sub TidyLeft(ByRef activePage As Microsoft.Office.Interop.Visio.Page, ByRef inShp As Visio.Shape)
		Try
			For Each shp In activePage.Shapes
				If shp.Name.Trim().StartsWith("EmptyDiamondL") Then
					shp.CellsU("LockDelete").ResultIU = 0
					shp.CellsU("LockGroup").ResultIU = 0 'shp.Protection.LockBegin.Value = True
					shp.CellsU("LockFromGroupFormat").ResultIU = 0
					'val.Ungroup()
					Debug.Write(shp.NameU + ", ")
					Dim selToDel As Visio.Selection
					selToDel = activePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty)
					Call selToDel.Select(shp, Visio.VisSelectArgs.visSelect)
					selToDel.Delete()
					Exit For
				End If
			Next shp
		Catch ex As Exception

		End Try
	End Sub

	'Tidy the right hand branch of the tree (flowchart) on stage 5 argument pattern.
	'This is super inefficient but we need to get and remove the shapes in the correct order.
	'We could have a single loop and find the shapes within one loop. However, this does not
	'guarantee the order and the deletion does not work. Also, some shapes are hidden and
	'onluy dscoverable when another 'higher' shape has gone.
	Public Shared Sub TidyRight(ByRef activePage As Microsoft.Office.Interop.Visio.Page, ByRef inShp As Visio.Shape)
		Try
			For Each shp In activePage.Shapes
				If shp.Name.Trim().StartsWith("J5_2Group") Then
					shp.CellsU("LockDelete").ResultIU = 0
					shp.CellsU("LockGroup").ResultIU = 0 'shp.Protection.LockBegin.Value = True
					shp.CellsU("LockFromGroupFormat").ResultIU = 0
					'val.Ungroup()
					Debug.Write(shp.NameU + ", ")
					Dim selToDel As Visio.Selection
					selToDel = activePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty)
					Call selToDel.Select(shp, Visio.VisSelectArgs.visSelect)
					selToDel.Delete()
					Exit For
				End If
			Next shp
		Catch errorThrown As Exception
			System.Diagnostics.Debug.WriteLine(errorThrown.Message)
		End Try
		Try
			For Each shp In activePage.Shapes
				If shp.Name.Trim().StartsWith("EmptyTriGroup") Then
					shp.CellsU("LockDelete").ResultIU = 0
					shp.CellsU("LockGroup").ResultIU = 0 'shp.Protection.LockBegin.Value = True
					shp.CellsU("LockFromGroupFormat").ResultIU = 0
					'val.Ungroup()
					Debug.Write(shp.NameU + ", ")
					Dim selToDel As Visio.Selection
					selToDel = activePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty)
					Call selToDel.Select(shp, Visio.VisSelectArgs.visSelect)
					selToDel.Delete()
					Exit For
				End If
			Next shp
		Catch errorThrown As Exception
			System.Diagnostics.Debug.WriteLine(errorThrown.Message)
		End Try
		Try
			For Each shp In activePage.Shapes
				If shp.Name.Trim().StartsWith("C5_2Group") Then
					shp.CellsU("LockDelete").ResultIU = 0
					shp.CellsU("LockGroup").ResultIU = 0 'shp.Protection.LockBegin.Value = True
					shp.CellsU("LockFromGroupFormat").ResultIU = 0
					'val.Ungroup()
					Debug.Write(shp.NameU + ", ")
					Dim selToDel As Visio.Selection
					selToDel = activePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty)
					Call selToDel.Select(shp, Visio.VisSelectArgs.visSelect)
					selToDel.Delete()
					Exit For
				End If
			Next shp
		Catch errorThrown As Exception
			System.Diagnostics.Debug.WriteLine(errorThrown.Message)
		End Try
		Try
			For Each shp In activePage.Shapes
				If shp.Name.Trim().StartsWith("JustificationGrouping") Then
					shp.CellsU("LockDelete").ResultIU = 0
					shp.CellsU("LockGroup").ResultIU = 0 'shp.Protection.LockBegin.Value = True
					shp.CellsU("LockFromGroupFormat").ResultIU = 0
					'val.Ungroup()
					Debug.Write(shp.NameU + ", ")
					Dim selToDel As Visio.Selection
					selToDel = activePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty)
					Call selToDel.Select(shp, Visio.VisSelectArgs.visSelect)
					selToDel.Delete()
					Exit For
				End If
			Next shp
		Catch errorThrown As Exception
			System.Diagnostics.Debug.WriteLine(errorThrown.Message)
		End Try
		Try
			For Each shp In activePage.Shapes
				If shp.Name.Trim().Equals("Ellipse") Then
					shp.CellsU("LockDelete").ResultIU = 0
					shp.CellsU("LockGroup").ResultIU = 0 'shp.Protection.LockBegin.Value = True
					shp.CellsU("LockFromGroupFormat").ResultIU = 0
					'val.Ungroup()
					Debug.Write(shp.NameU + ", ")
					Dim selToDel As Visio.Selection
					selToDel = activePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty)
					Call selToDel.Select(shp, Visio.VisSelectArgs.visSelect)
					selToDel.Delete()
					Exit For
				End If
			Next shp
		Catch errorThrown As Exception
			System.Diagnostics.Debug.WriteLine(errorThrown.Message)
		End Try
		Try
			For Each shp In activePage.Shapes
				If shp.Name.Trim().StartsWith("SheetEmptyTriangle") Then
					shp.CellsU("LockDelete").ResultIU = 0
					shp.CellsU("LockGroup").ResultIU = 0 'shp.Protection.LockBegin.Value = True
					shp.CellsU("LockFromGroupFormat").ResultIU = 0
					'val.Ungroup()
					Debug.Write(shp.NameU + ", ")
					Dim selToDel As Visio.Selection
					selToDel = activePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty)
					Call selToDel.Select(shp, Visio.VisSelectArgs.visSelect)
					selToDel.Delete()
					Exit For
				End If
			Next shp
		Catch errorThrown As Exception
			System.Diagnostics.Debug.WriteLine(errorThrown.Message)
		End Try
		Try
			For Each shp In activePage.Shapes
				If shp.Name.Trim().StartsWith("JustificationLetterJ") Then
					shp.CellsU("LockDelete").ResultIU = 0
					shp.CellsU("LockGroup").ResultIU = 0 'shp.Protection.LockBegin.Value = True
					shp.CellsU("LockFromGroupFormat").ResultIU = 0
					'val.Ungroup()
					Debug.Write(shp.NameU + ", ")
					Dim selToDel As Visio.Selection
					selToDel = activePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty)
					Call selToDel.Select(shp, Visio.VisSelectArgs.visSelect)
					selToDel.Delete()
					Exit For
				End If
			Next shp
		Catch errorThrown As Exception
			System.Diagnostics.Debug.WriteLine(errorThrown.Message)
		End Try
		Try
			For Each shp In activePage.Shapes
				If shp.Name.Trim().StartsWith("EmptyDiamondR") Then
					shp.CellsU("LockDelete").ResultIU = 0
					shp.CellsU("LockGroup").ResultIU = 0 'shp.Protection.LockBegin.Value = True
					shp.CellsU("LockFromGroupFormat").ResultIU = 0
					'val.Ungroup()
					Debug.Write(shp.NameU + ", ")
					Dim selToDel As Visio.Selection
					selToDel = activePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty)
					Call selToDel.Select(shp, Visio.VisSelectArgs.visSelect)
					selToDel.Delete()
					Exit For
				End If
			Next shp
		Catch errorThrown As Exception
			System.Diagnostics.Debug.WriteLine(errorThrown.Message)
		End Try



	End Sub

	'Post-order depth first search - tree traversal with deletion.
	'Removes the required branch of the tree (flowchart) on stage 5 argument pattern
	Public Shared Sub DepthFirstSearch(ByRef activePage As Microsoft.Office.Interop.Visio.Page, max As Integer, shapes As Visio.Shapes, shp As Visio.Shape)
		'Stack to store nodes during traversal
		Dim stack As Integer() = New Integer(max - 1) {}
		'Stack holding the nodes in post-order ready for deletion (LIFO deletion)
		Dim outStack As Visio.Shape() = New Visio.Shape(max - 1) {}
		'Array of nodes for indexing
		Dim arrVertices As Vertex() = New DFS.Vertex(max - 1) {}
		'Array of connections for doing the traversing
		Dim adjacencyMatrix As Integer(,) = New Integer(max - 1, max - 1) {}
		'Init array of connections
		For idx As Integer = 0 To max - 1
			For jdx As Integer = 0 To max - 1
				adjacencyMatrix(idx, jdx) = 0
			Next
		Next


		Dim stackTop As Integer = -1
		Dim outStackTop As Integer = -1
		'_vertexCount is number of nodes in array of nodes
		'Add new node to array of nodes
		_vertexCount = AddVertex(arrVertices, shp, 0)
		_nodeCount = 0
		'We've done the first nodes
		arrVertices(0).Visited = True
		'Add this node to the stack.
		'It will be the first node popped off the stack in the loop
		Dim t As Integer = stackTop
		stackTop = Push(stack, 0, t)

		While Not IsStackEmpty(stackTop) 'While nodes have not been processed (traversed)
			'Pop node off stack so we can process it
			Dim currentVertex As Integer = Pop(stack, stackTop)
			stackTop -= 1
			'Get the shape pointed to by index (currentVertex)
			shp = arrVertices(currentVertex).shape
			t = outStackTop
			'Add this shape to the list of shapes for deletion
			outStackTop = PushObject(outStack, shp, t)

			Dim shpIDs() As Integer
			'Get the 1-D shapes (arrows) connected to this shape (out connections) - if there are any
			shpIDs = shp.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing1D, "")
			'For i = 0 To UBound(shpIDs)
			'Debug.Print("Out 1d " + activePage.Shapes.ItemFromID(shpIDs(i)).NameU)
			'Next
			'Add the 2-D shapes connected to this shape (out connections)
			shpIDs = shpIDs.Union(shp.GluedShapes(Visio.VisGluedShapesFlags.visGluedShapesOutgoing2D, "")).Distinct().ToArray
			'MsgBox "Outgoing 2D shapes"
			'For i = 0 To UBound(shpIDs)
			'Debug.Print("Out all " + activePage.Shapes.ItemFromID(shpIDs(i)).Name)
			'Next
			'Process nodes - add them to the stack and mark them visited in the node array
			For i = 0 To (shpIDs.Length - 1)
				Dim v As Integer = _vertexCount
				_vertexCount = AddVertex(arrVertices, activePage.Shapes.ItemFromID(shpIDs(i)), v)
				t = stackTop
				stackTop = Push(stack, v, t)
				arrVertices(v).Visited = True
			Next
		End While

		Dim counter As Integer = outStackTop
		'Delete the nodes in order (LIFO) from the stack
		For i As Integer = 0 To counter - 1
			Dim val As Visio.Shape = PopObject(outStack, outStackTop)
			'Remove protection so we can delete them
			val.CellsU("LockDelete").ResultIU = 0
			val.CellsU("LockGroup").ResultIU = 0 'shp.Protection.LockBegin.Value = True
			val.CellsU("LockFromGroupFormat").ResultIU = 0
			outStackTop -= 1
			Debug.Write(val.NameU + ", ")
			Dim selToDel As Visio.Selection
			selToDel = activePage.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty)
			Call selToDel.Select(val, Visio.VisSelectArgs.visSelect)
			selToDel.Delete()
		Next
		Debug.Write(vbCrLf)
		'Tidy up
		For i As Integer = 0 To _vertexCount - 1
			'arrayVertices(i).Visited = False
			arrVertices(i).Visited = False
		Next
		_nodeCount = 0
		_vertexCount = 0
	End Sub

End Class
