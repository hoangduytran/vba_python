' macro_module.bas
' -----------------------------------------------------------
' Mục đích:
'   - Mở worksheet có tên "Planning"
'   - Tìm dòng chứa tiêu đề "Total" trong cột A. Nếu không tìm thấy,
'     sẽ tạo một dòng mới sau dòng dữ liệu cuối cùng.
'   - Tính tổng các giá trị số trong cột B (từ dòng 2 đến dòng trước dòng "Total")
'   - Ghi nhãn "Grand Total" vào cột A và tổng số vào cột B của dòng tổng.
'   - Thêm một cột mới bên phải bảng dữ liệu với tiêu đề "Percent" để
'     tính phần trăm của từng dòng dữ liệu so với tổng.
'   - Tạo biểu đồ cột hiển thị phần trăm, với nhãn trục X lấy từ cột A (giá trị "Month")
'
' Cách sử dụng:
'   1. Nhập file .bas này vào VBA Project của workbook.
'   2. Đảm bảo worksheet "Planning" có chứa dữ liệu phù hợp (ví dụ: cột A chứa "Month", cột B chứa số liệu).
'   3. Chạy macro "CreateGrandTotalAndChart".
' -----------------------------------------------------------

Sub CreateGrandTotalAndChart()
    Dim ws As Worksheet
    ' Gán worksheet "Planning" cho biến ws
    Set ws = ThisWorkbook.Worksheets("Planning")

    Dim lastRow As Long, lastCol As Long
    ' Xác định dòng cuối cùng có dữ liệu trong cột A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ' Xác định cột cuối cùng có dữ liệu trong dòng 1
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Tìm dòng có chứa từ "Total" trong cột A
    Dim totalRow As Long
    totalRow = 0  ' Khởi tạo với giá trị 0 (chưa tìm thấy dòng có "Total")
    Dim i As Long
    For i = 1 To lastRow
        If Trim(ws.Cells(i, 1).Value) = "Total" Then
            totalRow = i
            Exit For  ' Khi tìm thấy, thoát khỏi vòng lặp
        End If
    Next i

    ' Nếu không tìm thấy dòng "Total", thiết lập dòng tổng là ngay sau dòng dữ liệu cuối cùng
    If totalRow = 0 Then
        totalRow = lastRow + 1
    End If

    ' Giả định rằng dữ liệu số cần tính tổng nằm ở cột B (dataCol = 2)
    Dim dataCol As Long
    dataCol = 2

    ' Tính tổng các giá trị số từ dòng 2 đến dòng trước dòng "Total"
    Dim sumValue As Double
    sumValue = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(2, dataCol), ws.Cells(totalRow - 1, dataCol)))

    ' Ghi nhãn "Grand Total" vào cột A và tổng số vào cột B của dòng tổng
    ws.Cells(totalRow, 1).Value = "Grand Total"
    ws.Cells(totalRow, dataCol).Value = sumValue

    ' Thêm cột mới cho phần trăm bên phải cột cuối cùng hiện có
    Dim percentCol As Long
    percentCol = lastCol + 1
    ws.Cells(1, percentCol).Value = "Percent"

    ' Tính phần trăm cho mỗi dòng dữ liệu trong cột B (từ dòng 2 đến dòng trước dòng tổng)
    Dim currentVal As Double, percentVal As Double
    For i = 2 To totalRow - 1
        currentVal = ws.Cells(i, dataCol).Value
        ' Nếu tổng khác 0 thì tính phần trăm, nếu không thì gán giá trị 0
        If sumValue <> 0 Then
            percentVal = currentVal / sumValue
        Else
            percentVal = 0
        End If
        ' Ghi kết quả phần trăm vào ô tương ứng ở cột Percent
        ws.Cells(i, percentCol).Value = percentVal
        ' Định dạng ô hiển thị theo dạng phần trăm với 2 chữ số thập phân
        ws.Cells(i, percentCol).NumberFormat = "0.00%"
    Next i

    ' Tạo biểu đồ cột (Clustered Column Chart) để hiển thị phần trăm
    Dim chartObj As ChartObject
    Dim chartLeft As Double, chartTop As Double
    ' Xác định vị trí của biểu đồ: đặt bên cạnh bảng dữ liệu (có thể điều chỉnh lại vị trí theo ý muốn)
    chartLeft = ws.Cells(2, percentCol + 1).Left
    chartTop = ws.Cells(2, percentCol + 1).Top

    ' Thêm đối tượng biểu đồ vào worksheet với kích thước cụ thể
    Set chartObj = ws.ChartObjects.Add(Left:=chartLeft, Top:=chartTop, Width:=400, Height:=300)

    With chartObj.Chart
        .ChartType = xlColumnClustered  ' Đặt kiểu biểu đồ là cột nhóm (clustered column)
        ' Đặt nguồn dữ liệu cho biểu đồ:
        ' - Dòng 1 chứa tiêu đề "Percent"
        ' - Dòng 2 đến dòng trước dòng tổng chứa giá trị phần trăm
        .SetSourceData Source:=ws.Range(ws.Cells(1, percentCol), ws.Cells(totalRow - 1, percentCol))
        ' Đặt nhãn trục X cho biểu đồ lấy từ cột A (giá trị "Month") từ dòng 2 đến dòng trước dòng tổng
        .SeriesCollection(1).XValues = ws.Range(ws.Cells(2, 1), ws.Cells(totalRow - 1, 1))
        ' Thêm tiêu đề cho biểu đồ
        .HasTitle = True
        .ChartTitle.Text = "Percentage by Month"
    End With
End Sub
