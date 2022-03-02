Imports System.Xml

'https://visualsignals.typepad.co.uk/vislog/2012/03/building-visio-shapes-with-code.html
'Class to handle how many performance safety requirements there are and how many robustness
'safety requirements there are. The values are stored in a shape's shapesheet data.
'The shape is: Application.ActiveDocument.Pages.ItemU("Stage 5. Model Verification").Shapes("Document.238")
Public Class NumSRs
    Shared ReadOnly srsFile As String = "srsFile.xml"

    Public Shared Sub SetNumPerformanceSRs(val As Integer)
        ThisAddIn.numPerformanceSRs = val
    End Sub
    Public Shared Sub SetNumRobustnessSRs(val As Integer)
        ThisAddIn.numRobustnessSRs = val
    End Sub

    Public Shared Sub SetNumPerformanceSRsSet(val As Boolean)
        ThisAddIn.numPerformanceSRsSet = val
    End Sub
    Public Shared Sub SetNumRobustnessSRsSet(val As Boolean)
        ThisAddIn.numRobustnessSRsSet = val
    End Sub

    Public Shared Sub SetNumPerformanceSRsToFile(val As Integer)
        ThisAddIn.numPerformanceSRs = val
        ThisAddIn.numPerformanceSRsSet = True
        WriteSRsToFile(True, False)
    End Sub
    Public Shared Sub SetNumRobustnessSRsToFile(val As Integer)
        ThisAddIn.numRobustnessSRs = val
        ThisAddIn.numRobustnessSRsSet = True
        WriteSRsToFile(False, True)
    End Sub

    Public Shared Function GetNumPerformanceSRs() As Integer
        Return ThisAddIn.numPerformanceSRs
    End Function
    Public Shared Function GetNumRobustnessSRs() As Integer
        Return ThisAddIn.numRobustnessSRs
    End Function

    Public Shared Function GetNumPerformanceSRsSet() As Boolean
        Return ThisAddIn.numPerformanceSRsSet
    End Function
    Public Shared Function GetNumRobustnessSRsSet() As Boolean
        Return ThisAddIn.numRobustnessSRsSet
    End Function

    Public Shared Sub WriteSRsToFile(ByVal perfSRsSet As Boolean, ByVal robSRsSet As Boolean)
        'http://www.functionx.com/vb/xml/Lesson07.htm
        'http://vb.net-informations.com/xml/create-xml-vb.net.htm
        'Dim visFile As String = Replace(Globals.ThisAddIn.Application.ActiveDocument.Name, ".", "_")
        'Dim filename As String = visFile + "_" + srsFile
        'Dim writer As New XmlTextWriter(filename, System.Text.Encoding.UTF8)
        'writer.WriteStartDocument(True)
        'writer.Formatting = Formatting.Indented
        'writer.Indentation = 2
        'writer.WriteStartElement("SRs")
        'writer.WriteStartElement("performanceSRs")
        'writer.WriteString(GetNumPerformanceSRs().ToString)
        'writer.WriteEndElement()
        'writer.WriteStartElement("robustnessSRs")
        'writer.WriteString(GetNumRobustnessSRs().ToString)
        'writer.WriteEndElement()
        'writer.WriteEndElement()
        'writer.WriteEndDocument()
        'writer.Close()

        'http://visguy.com/vgforum/index.php?topic=4687.0
        'https://www.vbforums.com/showthread.php?398318-RESOLVED-Visio-automation
        Dim srShape As Visio.Shape
        Try
            srShape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU("Stage 5. Model Verification").Shapes("Document.238")
            Const SEC% = Visio.VisSectionIndices.visSectionUser

            If Not srShape.CellExists("User.perfSRs", 0) Then
                srShape.AddNamedRow(SEC%, "perfSRs", 0)
            End If
            Dim vCell = srShape.Cells("User.perfSRs")
            vCell.Formula = GetNumPerformanceSRs()

            If Not srShape.CellExists("User.robSRs", 0) Then
                srShape.AddNamedRow(SEC%, "robSRs", 0)
            End If
            vCell = srShape.Cells("User.robSRs")
            vCell.Formula = GetNumRobustnessSRs()

            Dim b As Boolean

            If Not srShape.CellExists("User.perfSRsSet", 0) Then
                srShape.AddNamedRow(SEC%, "perfSRsSet", 0)
            End If
            vCell = srShape.Cells("User.perfSRsSet")
            b = vCell.Formula
            If b = False Then
                vCell.Formula = perfSRsSet
            End If

            If Not srShape.CellExists("User.robSRsSet", 0) Then
                srShape.AddNamedRow(SEC%, "robSRsSet", 0)
            End If
            vCell = srShape.Cells("User.robSRsSet")
            b = vCell.Formula
            If b = False Then
                vCell.Formula = robSRsSet
            End If

        Catch ex As Exception
            'Shape is missing.
            '<TODO> Probably should do something here
        End Try

    End Sub

    Public Shared Sub ReadSRsFromFile()
        ''https://www.codeproject.com/Articles/4826/XML-File-Parsing-in-VB-NET
        'Dim visFile As String = Replace(Globals.ThisAddIn.Application.ActiveDocument.Name, ".", "_")
        'Dim filename As String = visFile + "_" + srsFile 'Globals.ThisAddIn.Application.ActiveDocument.Name ' Microsoft.Office.Interop.Visio.Document.Name + srsFile
        'Dim m_xmlr As XmlTextReader = Nothing
        ''Create the XML Reader
        'Try
        '    m_xmlr = New XmlTextReader(filename)
        '    'Disable whitespace so that you don't have to read over whitespaces
        '    m_xmlr.WhitespaceHandling = WhitespaceHandling.None
        '    'read the xml declaration and advance to SRs tag
        '    m_xmlr.Read()
        '    'read the SRs tag
        '    m_xmlr.Read()
        '    'Get the firstName Element Value
        '    m_xmlr.Read()
        '    Dim performanceSRsValue = m_xmlr.ReadElementString("performanceSRs")
        '    'Get the lastName Element Value
        '    Dim robustnessSRsValue = m_xmlr.ReadElementString("robustnessSRs")
        '    'close the reader
        '    m_xmlr.Close()
        '    'Write Results to the variables
        '    SetNumPerformanceSRs(Convert.ToInt32(performanceSRsValue))
        '    SetNumRobustnessSRs(Convert.ToInt32(robustnessSRsValue))
        'Catch ex As Exception
        '    m_xmlr.Close()
        '    SetNumPerformanceSRs(0)
        '    SetNumRobustnessSRs(0)
        'End Try
        Try
            Dim srShape As Visio.Shape
            srShape = Globals.ThisAddIn.Application.ActiveDocument.Pages.ItemU("Stage 5. Model Verification").Shapes("Document.238")
            Dim n As Integer

            n = 0
            If srShape.CellExists("User.perfSRs", 0) Then
                Dim vCell = srShape.Cells("User.perfSRs")
                n = CInt(Int(vCell.Formula))
            End If
            SetNumPerformanceSRs(n)

            n = 0
            If srShape.CellExists("User.robSRs", 0) Then
                Dim vCell = srShape.Cells("User.robSRs")
                n = CInt(Int(vCell.Formula))
            End If
            SetNumRobustnessSRs(n)

            Dim b As Boolean = False
            If srShape.CellExists("User.perfSRsSet", 0) Then
                Dim vCell = srShape.Cells("User.perfSRsSet")
                b = vCell.Formula
            End If
            SetNumPerformanceSRsSet(b)

            b = False
            If srShape.CellExists("User.robSRsSet", 0) Then
                Dim vCell = srShape.Cells("User.robSRsSet")
                b = vCell.Formula
            End If
            SetNumRobustnessSRsSet(b)
        Catch ex As Exception
            'Shape is missing.
            '<TODO> Probably should do something here
            SetNumPerformanceSRs(0)
            SetNumRobustnessSRs(0)
            SetNumPerformanceSRsSet(False)
            SetNumRobustnessSRsSet(False)
        End Try

    End Sub
End Class
