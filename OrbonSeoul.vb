Imports Excel = Microsoft.Office.Interop.Excel

Public Class OrbonSeoul
    Dim misValue As Object = System.Reflection.Missing.Value
    Public Sub getShippingList(ByVal path As String)
        Dim X As Object
        X = CreateObject("Excel.Application")

        X.Workbooks.Open(openFile())
        X.Calculation = Excel.XlCalculation.xlCalculationManual
        X.ScreenUpdating = False
        X.EnableEvents = False

        Dim B As Excel.Workbook
        Dim S As Excel.Worksheet
        Dim R As Excel.Range
        Dim Filename As String
        Dim savePath As String
        Dim watch As Stopwatch = Stopwatch.StartNew()

        B = X.ActiveWorkbook
        S = X.Worksheets(1)
        R = S.Range(S.Cells(1, 1), S.Cells(65000, 250))
        R.Columns.AutoFit()
        Dim rowCount As Integer = S.Range("A1", S.Range("A1").End(4)).Rows.Count
        If rowCount < 2 Or S.Cells(1, 1).Value <> "DLVR_CMPN_NM" Then
            MsgBox("잘못된 파일을 선택하였습니다.")
            B.Close(False)
            B = Nothing
            X.Quit()
            X = Nothing
            Return
        End If

        S.Range("A:A, C:I, K:L, O:U, W:W, Z:Z, AG:AY").EntireColumn.Delete()
        S.Cells(1, 7).Value = "BOX_SIZE"

        S.Range("H:H").NumberFormat = "@"
        S.Columns(10).Insert()
        S.Range("B:B").Copy()
        S.Paste(S.Range("J:J"))
        S.Columns(2).Delete
        S.Cells(1, 7).Value = "LOT_NO"

        Dim i As Integer
        For i = 2 To rowCount
            S.Cells(i, 7).Value = Right(S.Cells(i, 7).Value, 1) & "/" & Left(S.Cells(i, 7).Value, 1)
        Next

        S.Range("A1:M" & rowCount).Font.Size = 10
        S.Range("A1:M" & rowCount).RowHeight = 15.6
        S.Range("A1:M" & rowCount).HorizontalAlignment = Excel.Constants.xlCenter
        S.Range("A1:A1").EntireRow.Interior.Color = RGB(231, 230, 230)
        S.Range("A1:A1").ColumnWidth = 14
        S.Range("B1:B1").ColumnWidth = 7.5
        S.Range("C1:C1").ColumnWidth = 19
        S.Range("D1:D1").ColumnWidth = 8.8
        S.Range("E1:E1").ColumnWidth = 7.5
        S.Range("F1:F1").ColumnWidth = 8
        S.Range("G1:G1").ColumnWidth = 6
        S.Range("H1:H1").ColumnWidth = 14.4
        S.Range("I1:I1").ColumnWidth = 6.2
        S.Range("J1:J1").ColumnWidth = 6.8
        S.Range("K1:K1").ColumnWidth = 13.4
        S.Range("L1:L1").ColumnWidth = 9.2
        S.Range("M1:M1").ColumnWidth = 9.2
        For i = 2 To rowCount
            If S.Cells(i, 5).Value > 7 Then
                S.Cells(i, 6).Value = "대"
            ElseIf S.Cells(i, 5).Value > 3 Then
                S.Cells(i, 6).Value = "중"
            Else
                S.Cells(i, 6).Value = "소"
            End If
        Next
        S.Range("A1:A" & rowCount).Replace("_학교", "")
        S.Range("A1:A" & rowCount).Replace("1", "")

        S.Name = "전체"

        Dim misValue As Object = System.Reflection.Missing.Value
        Dim worksheets As Excel.Sheets = B.Worksheets

        Dim sGarak As Excel.Worksheet
        Dim sGarak2 As Excel.Worksheet
        Dim sGangseo As Excel.Worksheet
        Dim sGangseo2 As Excel.Worksheet
        Dim tmp As Excel.Worksheet

        tmp = worksheets.Add(worksheets(1), Type.Missing, Type.Missing, Type.Missing)
        tmp.Range("A1").Resize(1, S.Columns.Count).Value = S.Rows(1).Value

        If X.WorksheetFunction.CountIf(S.Range("I:I"), "2") > 0 Then
            tmp.Cells(2, 9).Value = 2
            sGarak = worksheets.Add(worksheets(1), Type.Missing, Type.Missing, Type.Missing)
            sGarak.Name = "가락"
            S.Range("A1:M" & rowCount).AdvancedFilter(Excel.XlFilterAction.xlFilterCopy, tmp.Range("A1").CurrentRegion, sGarak.Range("A1")) '가락
        End If

        If X.WorksheetFunction.CountIf(S.Range("I:I"), "84") > 0 Then
            tmp.Cells(2, 9).Value = 84
            sGarak2 = worksheets.Add(worksheets(1), Type.Missing, Type.Missing, Type.Missing)
            sGarak2.Name = "가락2층"
            S.Range("A1:M" & rowCount).AdvancedFilter(Excel.XlFilterAction.xlFilterCopy, tmp.Range("A1").CurrentRegion, sGarak2.Range("A1")) '가락2층
        End If

        If X.WorksheetFunction.CountIf(S.Range("I:I"), "1") > 0 Then
            tmp.Cells(2, 9).Value = 1
            sGangseo = worksheets.Add(worksheets(1), Type.Missing, Type.Missing, Type.Missing)
            sGangseo.Name = "강서"
            S.Range("A1:M" & rowCount).AdvancedFilter(Excel.XlFilterAction.xlFilterCopy, tmp.Range("A1").CurrentRegion, sGangseo.Range("A1")) '강서
        End If

        If X.WorksheetFunction.CountIf(S.Range("I:I"), "62") > 0 Then
            tmp.Cells(2, 9).Value = 62
            sGangseo2 = worksheets.Add(worksheets(1), Type.Missing, Type.Missing, Type.Missing)
            sGangseo2.Name = "강서지하"
            S.Range("A1:M" & rowCount).AdvancedFilter(Excel.XlFilterAction.xlFilterCopy, tmp.Range("A1").CurrentRegion, sGangseo2.Range("A1")) '강서지하
        End If

        X.DisplayAlerts = False
        tmp.Delete()
        X.DisplayAlerts = True

        Dim totalBox As Integer = rowCount - 1


        Dim j As Integer

        For i = 1 To worksheets.Count - 1
            worksheets(i).Range("D:D, F:F, G:G, I:I, J:J").EntireColumn.Delete()
            worksheets(i).Columns(4).Insert()
            worksheets(i).Range("A:A").Copy()
            worksheets(i).Paste(worksheets(i).Range("D:D"))
            worksheets(i).Columns(1).Delete
            worksheets(i).Cells(1, 1) = "NO"
            worksheets(i).Cells(1, 2) = "상품명"
            worksheets(i).Cells(1, 3) = "배송업체"
            worksheets(i).Cells(1, 4) = "중량"
            worksheets(i).Cells(1, 5) = "학교"
            worksheets(i).Cells(1, 6) = "좌표"
            worksheets(i).Cells(1, 7) = "제조일"
            worksheets(i).Cells(1, 8) = "유통기한"
            rowCount = worksheets(i).Range("A1", worksheets(i).Range("A1").End(4)).Rows.Count

            For j = 2 To rowCount
                Select Case worksheets(i).Cells(j, 1).Value
                    Case EProducts.SOYBEAN

                    Case EProducts.SOYBEAN_CA
                        worksheets(i).Cells(j, 2).Interior.Color = RGB(189, 215, 238)
                    Case EProducts.SOYBEAN_CHILD
                        worksheets(i).Cells(j, 2).Interior.Color = RGB(255, 255, 0)
                    Case EProducts.SOYBEAN_HEAD_CUT
                        worksheets(i).Cells(j, 2).Interior.Color = RGB(255, 192, 0)
                    Case EProducts.MUNGBEAN
                        worksheets(i).Cells(j, 2).Interior.Color = RGB(198, 224, 180)
                    Case Else
                        Debug.Assert(False)
                End Select
                worksheets(i).Cells(j, 1).Value = j - 1
            Next

            worksheets(i).Range("A1:H" & rowCount).Font.Size = 9
            worksheets(i).Range("A1:H" & rowCount).RowHeight = 17.3
            worksheets(i).Range("B:B").HorizontalAlignment = Excel.Constants.xlLeft
            worksheets(i).Range("C:C").HorizontalAlignment = Excel.Constants.xlLeft
            worksheets(i).Range("E:E").HorizontalAlignment = Excel.Constants.xlLeft
            worksheets(i).Range("A1:A1").ColumnWidth = 3.9
            worksheets(i).Range("B1:B1").ColumnWidth = 13.5
            worksheets(i).Range("C1:C1").ColumnWidth = 10.9
            worksheets(i).Range("D1:D1").ColumnWidth = 5.2
            worksheets(i).Range("E1:E1").ColumnWidth = 8.3
            worksheets(i).Range("F1:F1").ColumnWidth = 11.9
            worksheets(i).Range("G1:G1").ColumnWidth = 10.5
            worksheets(i).Range("H1:H1").ColumnWidth = 9.3
            worksheets(i).Range("B" & rowCount + 1 & ":D" & rowCount + 1).Font.Size = 9
            worksheets(i).Range("B" & rowCount + 1 & ":D" & rowCount + 1).HorizontalAlignment = Excel.Constants.xlCenter
            worksheets(i).Range("B" & rowCount + 1 & ":D" & rowCount + 1).Font.FontStyle = "Bold"
            worksheets(i).Range("B" & rowCount + 1 & ":D" & rowCount + 1).Interior.Color = RGB(208, 206, 206)
            worksheets(i).Cells(rowCount + 1, 2) = "계"
            worksheets(i).Cells(rowCount + 1, 4) = X.Sum(worksheets(i).Range("D2:D" & rowCount))
            worksheets(i).Range("A1:H1").HorizontalAlignment = Excel.Constants.xlCenter
            worksheets(i).Range("A1:H1").EntireRow.Interior.Color = RGB(231, 230, 230)
            With worksheets(i).PageSetup
                .PrintTitleRows = "$1:$1"
                .LeftHeader = "&12" & worksheets(i).name
                .CenterHeader = "&12" & " 센터 박스 수 : " & rowCount - 1
                .RightHeader = "&12출하일 : " & Left(worksheets(i).Cells(2, 7).Value, 5) & "-" & Left(Right(worksheets(i).Cells(2, 7).Value, 5), 2) & "-" & Right(worksheets(i).Cells(2, 7).Value, 2)
                .LeftFooter = "&12출력일 : " & Format(DateTime.Now, "yyyy-MM-dd HH:mm tt")
                .CenterFooter = "&12금일 총 박스수 : " & totalBox
                .RightFooter = "&P/&N"
                .LeftMargin = X.CentimetersToPoints(1.3)
                .RightMargin = X.CentimetersToPoints(1.3)
                .TopMargin = X.CentimetersToPoints(2.5)
                .BottomMargin = X.CentimetersToPoints(2.5)
                .HeaderMargin = X.CentimetersToPoints(1.5)
                .FooterMargin = X.CentimetersToPoints(1.5)
            End With
            worksheets(i).Range("A1:H" & rowCount + 1).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic)
            worksheets(i).Range("A1:H" & rowCount + 1).Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlDash
            worksheets(i).Range("A1:H" & rowCount + 1).Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlDash
            worksheets(i).Range("A1:H1").Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.Constants.xlNone
            worksheets(i).Range("A1:H1").Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.Constants.xlNone



        Next
        Dim S_summary As Excel.Worksheet
        S_summary = worksheets.Add(worksheets(1), Type.Missing, Type.Missing, Type.Missing)
        S_summary.Name = "집계"

        S_summary.Cells(1, 1).Value = "품목"
        S_summary.Cells(1, 2).Value = "중량(kg)"
        S_summary.Cells(1, 3).Value = "박스수"
        S_summary.Cells(1, 4).Value = "대"
        S_summary.Cells(1, 5).Value = "중"
        S_summary.Cells(1, 6).Value = "소"
        S_summary.Cells(2, 1).Value = "두절콩나물"
        S_summary.Cells(3, 1).Value = "새싹콩나물(어린콩나물)"
        S_summary.Cells(4, 1).Value = "숙주나물(손질)"
        S_summary.Cells(5, 1).Value = "콩나물(손질)"
        S_summary.Cells(6, 1).Value = "콩나물(칼슘콩나물)"
        S_summary.Cells(7, 1).Value = "계"

        S_summary.Range("A1:F1").Interior.Color = RGB(217, 217, 217)
        S_summary.Range("A2:A6").Interior.Color = RGB(242, 242, 242)
        S_summary.Range("A7:F7").Interior.Color = RGB(217, 217, 217)

        S_summary.Cells(9, 1).Value = "센터"
        S_summary.Cells(9, 2).Value = "중량"
        S_summary.Cells(9, 3).Value = "박스수"
        S_summary.Cells(10, 1).Value = "가락"
        S_summary.Cells(11, 1).Value = "가락2층"
        S_summary.Cells(12, 1).Value = "강서"
        S_summary.Cells(13, 1).Value = "강서지하"
        S_summary.Cells(14, 1).Value = "계"

        S_summary.Range("A9:C9").Interior.Color = RGB(217, 217, 217)
        S_summary.Range("A10:A13").Interior.Color = RGB(242, 242, 242)
        S_summary.Range("A14:C14").Interior.Color = RGB(217, 217, 217)

        S_summary.Cells(16, 1).Value = "배송업체명"
        S_summary.Cells(16, 2).Value = "합계 : 발주량"
        S_summary.Cells(16, 3).Value = "합계 : 박스수"
        S_summary.Cells(17, 1).Value = "대한"
        S_summary.Cells(18, 1).Value = "동선"
        S_summary.Cells(19, 1).Value = "모닝"
        S_summary.Cells(20, 1).Value = "미림"
        S_summary.Cells(21, 1).Value = "보상"
        S_summary.Cells(22, 1).Value = "보은농산"
        S_summary.Cells(23, 1).Value = "삼성"
        S_summary.Cells(24, 1).Value = "상촌"
        S_summary.Cells(25, 1).Value = "서부"
        S_summary.Cells(26, 1).Value = "수산"
        S_summary.Cells(27, 1).Value = "아랑"
        S_summary.Cells(28, 1).Value = "양양"
        S_summary.Cells(29, 1).Value = "엘케이푸드"
        S_summary.Cells(30, 1).Value = "이조은유기농"
        S_summary.Cells(31, 1).Value = "자연"
        S_summary.Cells(32, 1).Value = "정성"
        S_summary.Cells(33, 1).Value = "초원에프에스"
        S_summary.Cells(34, 1).Value = "하나"
        S_summary.Cells(35, 1).Value = "해움터"
        S_summary.Cells(36, 1).Value = "현대농산"
        S_summary.Cells(37, 1).Value = "현진그린"
        S_summary.Cells(38, 1).Value = "총합계"

        With S_summary.Range("A16:C16")
            .Interior.Color = RGB(191, 191, 191)
            .Font.Size = 11
            .Font.FontStyle = "Bold"
            .RowHeight = 19
        End With

        With S_summary.Range("A38:C38")
            .Interior.Color = RGB(191, 191, 191)
            .Font.Size = 11
            .Font.FontStyle = "Bold"
            .RowHeight = 19
        End With

        With S_summary.Range("A1:F14")
            .Font.Size = 10
            .RowHeight = 19
        End With

        With S_summary.Range("A2:C6")
            .Font.Size = 9
        End With

        With S_summary.Range("A1:F38")
            .HorizontalAlignment = Excel.Constants.xlCenter
        End With


        S_summary.Range("A1:A1").ColumnWidth = 22.6
        S_summary.Range("A17:C37").RowHeight = 15
        S_summary.Range("A17:A37").Interior.Color = RGB(242, 242, 242)
        S_summary.Range("B1:B1").ColumnWidth = 12
        S_summary.Range("B1:B1").NumberFormatLocal = "0.00"
        S_summary.Range("C1:C1").ColumnWidth = 12

        rowCount = S.Range("A1", S.Range("A1").End(4)).Rows.Count
        Dim summaryRow As Integer
        For i = 2 To rowCount
            Select Case S.Cells(i, 2).Value
                Case EProducts.SOYBEAN
                    summaryRow = 5
                Case EProducts.SOYBEAN_CA
                    summaryRow = 6
                Case EProducts.SOYBEAN_CHILD
                    summaryRow = 3
                Case EProducts.SOYBEAN_HEAD_CUT
                    summaryRow = 2
                Case EProducts.MUNGBEAN
                    summaryRow = 4
            End Select
            '중량 계산
            S_summary.Cells(summaryRow, 2).Value = S_summary.Cells(summaryRow, 2).Value + S.Cells(i, 5).Value
            '박스수 계산
            S_summary.Cells(summaryRow, 3).Value = S_summary.Cells(summaryRow, 3).Value + 1
            '대중소 계산
            If S.Cells(i, 6).Value = "대" Then
                S_summary.Cells(summaryRow, 4).Value = S_summary.Cells(summaryRow, 4).Value + 1
            ElseIf S.Cells(i, 6).Value = "중" Then
                S_summary.Cells(summaryRow, 5).Value = S_summary.Cells(summaryRow, 5).Value + 1
            Else
                S_summary.Cells(summaryRow, 6).Value = S_summary.Cells(summaryRow, 6).Value + 1
            End If

            '센터 중량 계산
            Select Case S.Cells(i, 9).Value
                Case EStrg.GARAK
                    summaryRow = 10
                Case EStrg.GARAK_FLOOR2
                    summaryRow = 11
                Case EStrg.GANGSEO
                    summaryRow = 12
                Case EStrg.GANGSEO_UNDER_FLOOR
                    summaryRow = 13
            End Select
            S_summary.Cells(summaryRow, 2).Value = S_summary.Cells(summaryRow, 2).Value + S.Cells(i, 5).Value
            '센터 박스수 계산
            S_summary.Cells(summaryRow, 3).Value = S_summary.Cells(summaryRow, 3).Value + 1
            '업체별 박스수 계산
            Dim str = S.Cells(i, 1).Value
            Dim matchIndex As Integer = X.WorksheetFunction.Match(S.Cells(i, 1), S_summary.Range("A1:A37"), 0)
            S_summary.Cells(matchIndex, 3).Value = S_summary.Cells(matchIndex, 3).Value + 1
            '업체별 중량 계산
            S_summary.Cells(matchIndex, 2).Value = S_summary.Cells(matchIndex, 2).Value + S.Cells(i, 5).Value
        Next

        For i = 2 To 6
            S_summary.Cells(7, i).Value = S_summary.Cells(2, i).Value + S_summary.Cells(3, i).Value + S_summary.Cells(4, i).Value + S_summary.Cells(5, i).Value + S_summary.Cells(6, i).Value
        Next

        For i = 2 To 3
            S_summary.Cells(14, i).Value = S_summary.Cells(10, i).Value + S_summary.Cells(11, i).Value + S_summary.Cells(12, i).Value + S_summary.Cells(13, i).Value
        Next
        S_summary.Cells(38, 2).Value = X.Sum(S_summary.Range("B17:B37"))
        S_summary.Cells(38, 3).Value = X.Sum(S_summary.Range("C17:C37"))
        Filename = S.Cells(2, 12).Value & " (" & Now.ToString("yyyy/MM/dd/hh/mm/ss") & ")"
        savePath = TTT(path)
        path = savePath & Filename & ".xlsx"
        X.DisplayAlerts = False
        B.SaveAs(path, 51)
        X.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        X.ScreenUpdating = True
        X.EnableEvents = True
        X.Visible = True
        'B.Close()
        'B = Nothing
        watch.Stop()
        'X.Quit()
        'X = Nothing
    End Sub
    Private Function openFile()
        Dim ofd As New OpenFileDialog()
        Dim res As DialogResult
        With ofd
            .FileName = ""
            .InitialDirectory = "C:\Users\" & SystemInformation.UserName & "\Desktop"
            .Filter = "TXT(*.txt)|*.txt"
            .Title = "올본 출하 리스트 선택"
            .Multiselect = False
            res = .ShowDialog
            openFile = ofd.FileName
        End With
    End Function


    Private Function TTT(ByVal strPath As String) As String

        Dim path As String
        path = strPath & "\"

        If Len(Dir(path, vbDirectory)) <= 0 Then
            MkDir(path)
        End If

        TTT = path
    End Function
End Class
