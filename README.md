# VBA-code-to-automate-values-extraction
VBA code to automate values extraction
The code extracts values meeting a specific criteria, from multiple pivot tables which are saved in a separate workbook.
The structure of the VBA code is dynamic such that it extracts values based on multiple criterias and therefore does not require any ongoing adjustments.



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

 

    filepath = \\svrau530csm00\DOITribe\1. GoBiz\8. Data & Reporting\DemystData Reporting\

    filename = Dir(filepath & "Report -*.xlsx")

 

    Set WbDest = ThisWorkbook

    Set WsDest = WbDest.Worksheets("Start Here")

    Set WbSource = Workbooks.Open(filepath & filename)

    Set Wssource = WbSource.Worksheets("LH Summary")

    Set pivot = Wssource.PivotTables("DrawnBySettlement")

    Set pivot2 = Wssource.PivotTables("DrawnBySubmitted")

 

    refDate = WsDest.Range("B3").value

    refMonth = Format(refDate, "Mmm")

    refYear = Format(refDate, "yyyy")

 

    With pivot

        countValue = IIf(IsError(.GetPivotData("Count", "Drawn Date", refMonth, "Years", refYear)), 0, .GetPivotData("Count", "Drawn Date", refMonth, "Years", refYear).value)

        drawnDollarValue = IIf(IsError(.GetPivotData("Drawn $", "Drawn Date", refMonth, "Years", refYear)), 0, .GetPivotData("Drawn $", "Drawn Date", refMonth, "Years", refYear).value)

    End With

 

    With pivot2

        countValue2 = IIf(IsError(.GetPivotData("Count of BBD APP #", "Date", refMonth, "Years2", refYear)), 0, .GetPivotData("Count of BBD APP #", "Date", refMonth, "Years2", refYear).value)

        drawnDollarValue2 = IIf(IsError(.GetPivotData("Drawn Amount", "Date", refMonth, "Years2", refYear)), 0, .GetPivotData("Drawn Amount", "Date", refMonth, "Years2", refYear).value)

    End With

 

    Set colrange = WsDest.Range("B8:P8")

 

    For Each cell In colrange

        If cell.value = refDate Then

            WsDest.Cells(40, cell.Column).value = drawnDollarValue

            WsDest.Cells(45, cell.Column).value = countValue

            WsDest.Cells(29, cell.Column).value = drawnDollarValue2

            WsDest.Cells(34, cell.Column).value = countValue2

        End If

    Next cell

 

    WbSource.Close
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

 

    filepath = \\Data & Reporting\Reporting\

    filename = Dir(filepath & "Report -*.xlsx")

 

    Set WbDest = ThisWorkbook

    Set WsDest = WbDest.Worksheets("Start Here")

    Set WbSource = Workbooks.Open(filepath & filename)

    Set Wssource = WbSource.Worksheets("LH Summary")

    Set pivot = Wssource.PivotTables("DrawnBySettlement")

    Set pivot2 = Wssource.PivotTables("DrawnBySubmitted")

 

    refDate = WsDest.Range("B3").value

    refMonth = Format(refDate, "Mmm")

    refYear = Format(refDate, "yyyy")

 

    With pivot

        countValue = IIf(IsError(.GetPivotData("Count", "Drawn Date", refMonth, "Years", refYear)), 0, .GetPivotData("Count", "Drawn Date", refMonth, "Years", refYear).value)

        drawnDollarValue = IIf(IsError(.GetPivotData("Drawn $", "Drawn Date", refMonth, "Years", refYear)), 0, .GetPivotData("Drawn $", "Drawn Date", refMonth, "Years", refYear).value)

    End With

 

    With pivot2

        countValue2 = IIf(IsError(.GetPivotData("Count of BBD APP #", "Date", refMonth, "Years2", refYear)), 0, .GetPivotData("Count of BBD APP #", "Date", refMonth, "Years2", refYear).value)

        drawnDollarValue2 = IIf(IsError(.GetPivotData("Drawn Amount", "Date", refMonth, "Years2", refYear)), 0, .GetPivotData("Drawn Amount", "Date", refMonth, "Years2", refYear).value)

    End With

 

    Set colrange = WsDest.Range("B8:P8")

 

    For Each cell In colrange

        If cell.value = refDate Then

            WsDest.Cells(40, cell.Column).value = drawnDollarValue

            WsDest.Cells(45, cell.Column).value = countValue

            WsDest.Cells(29, cell.Column).value = drawnDollarValue2

            WsDest.Cells(34, cell.Column).value = countValue2

        End If

    Next cell

 

    WbSource.Close
