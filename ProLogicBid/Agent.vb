Public Class Agent

    Public Sub InitiateBidProposal(ID As String)
        Dim startInfo As New ProcessStartInfo
        startInfo.FileName = ("C:\Users\darrenm\Desktop\ProLogicReportingApplication\ProLogicReportingApplication\bin\Debug\ProLogicReportingApplication.exe")
        startInfo.Arguments = ID
        Process.Start(startInfo)
    End Sub

    Public Event SendEvent(ByRef activityType As String, ByRef XMLIn As String)

    Public Function SendEventDelegate(activityType As String, XMLIn As String)
        RaiseEvent SendEvent(activityType, XMLIn)
        Return ""
    End Function

End Class
