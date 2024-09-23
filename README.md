# VBA-code-to-automate-values-extraction
VBA code to automate values extraction
The code extracts values from 2 pivot tables which are saved in a separate workbook.
The structure of the VBA code is dynamic such that it extracts values based on a date criteria and therefore does not require any adjustments.


Public Sub GetDrawnValues()

 

    Dim WbDest As Workbook

    Dim WbSource As Workbook

    Dim filepath As String

    Dim filename As String

    Dim WsDest As Worksheet

    Dim Wssource As Worksheet

    Dim refDate As Date

    Dim refMonth As String

    Dim refYear As String

    Dim pivot As PivotTable

    Dim pivot2 As PivotTable

    Dim countValue As Double

    Dim drawnDollarValue As Double

    Dim countValue2 As Double

    Dim drawnDollarValue2 As Double

    Dim colrange As Range

    Dim cell As Range

 

    filepath = \\Data Reporting\Test\

    filename = Dir(filepath & "Report -*.xlsx")

 

    Set WbDest = ThisWorkbook

    Set WsDest = WbDest.Worksheets("Start Here")

    Set WbSource = Workbooks.Open(filepath & filename)

    Set Wssource = WbSource.Worksheets("LH Summary")

    Set pivot = Wssource.PivotTables("DrawnBySettlement")

    Set pivot2 = Wssource.PivotTables("DrawnBySubmitted")

 

    refDate = WsDest.Range("B3").Value

    refMonth = Format(refDate, "Mmm")

    refYear = Format(refDate, "yyyy")

 

    With pivot

        countValue = .GetPivotData("Count", "Drawn Date", refMonth, "Years", refYear).Value

        drawnDollarValue = .GetPivotData("Drawn $", "Drawn Date", refMonth, "Years", refYear).Value

    End With

 

    With pivot2

        countValue2 = .GetPivotData("Count of BBD APP #", "Date", refMonth, "Years2", refYear).Value

        drawnDollarValue2 = .GetPivotData("Drawn Amount", "Date", refMonth, "Years2", refYear).Value

    End With

 

    Set colrange = WsDest.Range("B8:P8")

 

    For Each cell In colrange

        If cell.Value = refDate Then

            WsDest.Cells(40, cell.Column).Value = drawnDollarValue

            WsDest.Cells(45, cell.Column).Value = countValue

            WsDest.Cells(29, cell.Column).Value = drawnDollarValue2

            WsDest.Cells(34, cell.Column).Value = countValue2

        End If

    Next cell

 

    WbSource.Close

 

End Sub
