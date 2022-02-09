Imports System.Xml

'https://visualsignals.typepad.co.uk/vislog/2012/03/building-visio-shapes-with-code.html
Public Class NumSRs
    Shared ReadOnly srsFile As String = "srsFile.xml"

    Public Shared Sub SetNumPerformanceSRs(val As Integer)
        ThisAddIn.numPerformanceSRs = val
    End Sub
    Public Shared Sub SetNumRobustnessSRs(val As Integer)
        ThisAddIn.numRobustnessSRs = val
    End Sub

    Public Shared Sub SetNumPerformanceSRsToFile(val As Integer)
        ThisAddIn.numPerformanceSRs = val
        WriteSRsToFile()
    End Sub
    Public Shared Sub SetNumRobustnessSRsToFile(val As Integer)
        ThisAddIn.numRobustnessSRs = val
        WriteSRsToFile()
    End Sub

    Public Shared Function GetNumPerformanceSRs() As Integer
        Return ThisAddIn.numPerformanceSRs
    End Function
    Public Shared Function GetNumRobustnessSRs() As Integer
        Return ThisAddIn.numRobustnessSRs
    End Function
    Public Shared Sub WriteSRsToFile()
        'http://www.functionx.com/vb/xml/Lesson07.htm
        'http://vb.net-informations.com/xml/create-xml-vb.net.htm
        Dim visFile As String = Replace(Globals.ThisAddIn.Application.ActiveDocument.Name, ".", "_")
        Dim filename As String = visFile + "_" + srsFile
        Dim writer As New XmlTextWriter(filename, System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        writer.WriteStartElement("SRs")
        writer.WriteStartElement("performanceSRs")
        writer.WriteString(GetNumPerformanceSRs().ToString)
        writer.WriteEndElement()
        writer.WriteStartElement("robustnessSRs")
        writer.WriteString(GetNumRobustnessSRs().ToString)
        writer.WriteEndElement()
        writer.WriteEndElement()
        writer.WriteEndDocument()
        writer.Close()
    End Sub

    Public Shared Sub ReadSRsFromFile()
        'https://www.codeproject.com/Articles/4826/XML-File-Parsing-in-VB-NET
        Dim visFile As String = Replace(Globals.ThisAddIn.Application.ActiveDocument.Name, ".", "_")
        Dim filename As String = visFile + "_" + srsFile 'Globals.ThisAddIn.Application.ActiveDocument.Name ' Microsoft.Office.Interop.Visio.Document.Name + srsFile
        Dim m_xmlr As XmlTextReader = Nothing
        'Create the XML Reader
        Try
            m_xmlr = New XmlTextReader(filename)
            'Disable whitespace so that you don't have to read over whitespaces
            m_xmlr.WhitespaceHandling = WhitespaceHandling.None
            'read the xml declaration and advance to SRs tag
            m_xmlr.Read()
            'read the SRs tag
            m_xmlr.Read()
            'Get the firstName Element Value
            m_xmlr.Read()
            Dim performanceSRsValue = m_xmlr.ReadElementString("performanceSRs")
            'Get the lastName Element Value
            Dim robustnessSRsValue = m_xmlr.ReadElementString("robustnessSRs")
            'close the reader
            m_xmlr.Close()
            'Write Results to the variables
            SetNumPerformanceSRs(Convert.ToInt32(performanceSRsValue))
            SetNumRobustnessSRs(Convert.ToInt32(robustnessSRsValue))
        Catch ex As Exception
            m_xmlr.Close()
            SetNumPerformanceSRs(0)
            SetNumRobustnessSRs(0)
        End Try

    End Sub
End Class
