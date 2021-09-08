Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        MsgBox("testing", Title:="Test me!")
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
