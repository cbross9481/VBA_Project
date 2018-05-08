
Sub SampleData2()
    ' import file from Extract Data workbook
    
    
    Dim x As Integer, y As Integer, i As Integer, tb As ListObject, hidden As range
    
    Set tb = ActiveSheet.ListObjects("SampleTable")
    
    i = 1
    Do While i < tb.range.Rows.Count
    
        tb.DataBodyRange.Cells(i, tb.ListColumns("Unit Cost").Index) = WorksheetFunction.RandBetween(1, 5)
        tb.DataBodyRange.Cells(i, tb.ListColumns("Units").Index) = WorksheetFunction.RandBetween(1, 100)
        tb.DataBodyRange.Cells(i, tb.ListColumns("Total").Index) = tb.DataBodyRange.Cells(i, tb.ListColumns("Unit Cost").Index) * tb.DataBodyRange.Cells(i, tb.ListColumns("Units").Index)
        i = i + 1
    Loop
    
    Sort_Data tb
    FindLargestValues total, tb
    Sort_Data2 tb
    
End Sub

Sub FindLargestValues(total, tb)
    
    Dim tbl As ListObject, listing As ListObject
    Set tbl = ActiveSheet.ListObjects("tierTable")
    Set listing = ActiveSheet.ListObjects("Listing")
    
    i = 1
    Dim mylookupvalue As Integer, vtable As range, column As Long
    
    Do While i < listing.range.Rows.Count
        listing.DataBodyRange(i, 5).Value = tbl.DataBodyRange(1, i) * tb.DataBodyRange.Cells(i, tb.ListColumns("Total").Index)
        listing.DataBodyRange(i, 4).Value = tb.DataBodyRange.Cells(i, tb.ListColumns("Total").Index)
        listing.DataBodyRange(i, 2).Value = tb.DataBodyRange.Cells(i, tb.ListColumns("Rep").Index)
        listing.DataBodyRange(i, 3).Value = tb.DataBodyRange.Cells(i, tb.ListColumns("Item").Index)
        
        i = i + 1
    Loop

End Sub

Sub Sort_Data(tb)
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("SampleTable").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("SampleTable").Sort.SortFields. _
        Add Key:=range("SampleTable[Total]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("SampleTable").Sort.SortFields. _
        Add Key:=range("SampleTable[OrderDate]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").ListObjects("SampleTable").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub Sort_Data2(tb)

    With tb.Sort
        .SortFields.Clear
        .SortFields.Add _
        Key:=tb.ListColumns("OrderDate").range, _
        SortOn:=xlSortOnValues, _
        Order:=xlDescending, _
        DataOption:=xlSortNormal
        .Apply
    End With

End Sub
