Module check

    Sub Main()

        Dim larr_Args() As String
        larr_Args = Split(Command, " ")

        If larr_Args.Length <> 2 Then
            Console.WriteLine("Khong dung so tham so, can 2 tham so: tepExcelcuaSV tepExcelChuan")
            Exit Sub
        End If

        Dim oE As New OpenExcel

        If oE.check(larr_Args(0), larr_Args(1), 100, 3, 1, 6) Then
            Console.WriteLine("Dung")
        Else
            Console.WriteLine("Sai")
        End If

    End Sub

End Module
