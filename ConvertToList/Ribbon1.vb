Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub convertButton_Click(sender As Object, e As RibbonControlEventArgs) Handles convertButton.Click
        Dim xlApp As Excel.Application = Globals.ThisAddIn.Application
        Dim xlRange As Excel.Range = CType(xlApp.Selection, Excel.Range)

        Dim c As Excel.Range
        Dim thisRange As Excel.Range
        Dim delimiter = delimiterBox.Text

        If delimiter = "" Then
            delimiter = xlApp.InputBox("Please enter a delimiter to use")
        End If

        If xlRange IsNot Nothing Then
            If xlRange.Cells.Count = 1 Then
                thisRange = xlRange
            Else
                thisRange = xlRange.SpecialCells(Excel.XlCellType.xlCellTypeConstants)
            End If
            For Each c In thisRange
                Dim text() As String
                Dim bullet As String
                Dim html As String
                html = "<ul>"
                text = Split(c.Text, delimiter)

                For Each bullet In text
                    html += "<li>"
                    html += bullet
                    html += "</li>"
                Next
                html += "</ul>"
                c.Value2 = html
            Next
        End If
    End Sub
End Class
