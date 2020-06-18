Public Class OpenExcel

    Public Function check(ByVal StudentFile As String, ByVal OriginFile As String, ByVal dataRows As Integer, ByVal dataCols As Integer, ByVal resultPosRow As Integer, ByVal resultPosCol As Integer) As Boolean
        Dim xlApp As Excel.Application
        Dim xl As Excel.Workbook

        Dim x2App As Excel.Application
        Dim x2 As Excel.Workbook

        xlApp = GetObject("", "Excel.Application")    'xlApp = New Excel.Application Seems to cause a memory leak
        xlApp.Visible = False      'no display Excel Application
        xlApp.DisplayAlerts = False      'Don't show any message, like "SAVE?"

        xl = xlApp.Workbooks.Open(System.Environment.CurrentDirectory + "\" + OriginFile, , True)

        x2App = GetObject("", "Excel.Application")    'xlApp = New Excel.Application Seems to cause a memory leak
        x2App.Visible = False      'no display Excel Application
        x2App.DisplayAlerts = False      'Don't show any message, like "SAVE?"

        x2 = x2App.Workbooks.Open(System.Environment.CurrentDirectory + "\" + StudentFile, , True)

        Dim i, j As Integer

        For i = 1 To dataRows
            For j = 1 To dataCols
                x2.Sheets("Sheet1").Cells(i, j).MergeArea.Value() = xl.Sheets("Sheet1").Cells(i, j).MergeArea.Value()
            Next
        Next

        If x2.Sheets("Sheet1").Cells(resultPosRow, resultPosCol).MergeArea.Value() = xl.Sheets("Sheet1").Cells(resultPosRow, resultPosCol).MergeArea.Value() Then
            check = True
        Else
            check = False
        End If

        ' Dong du lieu
        xl.Application.Workbooks.Close()
        xl = Nothing
        xlApp = Nothing

        x2.Application.Workbooks.Close()
        x2 = Nothing
        x2App = Nothing

    End Function

End Class
