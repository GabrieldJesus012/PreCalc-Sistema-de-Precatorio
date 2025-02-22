Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim xlApp As Object
        Dim xlWb As Object
        Dim xlWbName As String
        Dim xlWbPath As String

        On Error Resume Next

        xlApp = CreateObject("Excel.Application")

        xlWbName = "PreCalc.xlsm"
        xlWbPath = Application.StartupPath

        Me.Hide()

        xlWb = xlApp.workbooks.open(xlWbPath & "\" & xlWbName)


        xlApp = Nothing
        xlWbName = Nothing
        xlWbPath = Nothing
        xlWb = Nothing

        Application.Exit()
    End Sub
End Class
