Imports Microsoft.Office.Interop.Excel
Imports System.Collections.Generic
Imports System.Runtime.InteropServices

Public Class ExcelHelper

    Public Shared Sub ExportToExcel(data As List(Of ObjectData), filePath As String, sheetName As String)
        Dim excelApp As Application = Nothing
        Dim workbook As Workbook = Nothing
        Dim worksheet As Worksheet = Nothing

        Try
            Logger.Log($"엑셀 내보내기 시작: {filePath}")

            excelApp = New Application()
            excelApp.Visible = False
            excelApp.DisplayAlerts = False

            workbook = excelApp.Workbooks.Add()
            worksheet = CType(workbook.Worksheets(1), Worksheet)
            worksheet.Name = sheetName

            ' 헤더 작성
            Dim headers = {"LinkFileName", "Id", "ElementName", "CategoryName", "LocationX", "LocationY", "LocationZ",
                          "Grid", "SB_NAME", "SB_FIELD", "SB_FL", "SB_BLDG", "Stud Number", "Column Number",
                          "EXCLUSION", "CLASS", "SAMOO-Other_1", "SAMOO-Other_2", "SAMOO-Other_3",
                          "SECC-Other_3", "SECC-Other_4"}

            For col As Integer = 0 To headers.Length - 1
                worksheet.Cells(1, col + 1) = headers(col)
            Next

            ' 데이터 작성
            For row As Integer = 0 To data.Count - 1
                Dim obj = data(row)
                worksheet.Cells(row + 2, 1) = obj.LinkFileName
                worksheet.Cells(row + 2, 2) = obj.ElementId
                worksheet.Cells(row + 2, 3) = obj.ElementName
                worksheet.Cells(row + 2, 4) = obj.CategoryName
                worksheet.Cells(row + 2, 5) = obj.LocationX
                worksheet.Cells(row + 2, 6) = obj.LocationY
                worksheet.Cells(row + 2, 7) = obj.LocationZ
                worksheet.Cells(row + 2, 8) = obj.Grid
                worksheet.Cells(row + 2, 9) = obj.SB_NAME
                worksheet.Cells(row + 2, 10) = obj.SB_FIELD
                worksheet.Cells(row + 2, 11) = obj.SB_FL
                worksheet.Cells(row + 2, 12) = obj.SB_BLDG
                worksheet.Cells(row + 2, 13) = obj.StudNumber
                worksheet.Cells(row + 2, 14) = obj.ColumnNumber
                worksheet.Cells(row + 2, 15) = obj.EXCLUSION
                worksheet.Cells(row + 2, 16) = obj.CLASS
                worksheet.Cells(row + 2, 17) = obj.SAMOO_Other_1
                worksheet.Cells(row + 2, 18) = obj.SAMOO_Other_2
                worksheet.Cells(row + 2, 19) = obj.SAMOO_Other_3
                worksheet.Cells(row + 2, 20) = obj.SECC_Other_3
                worksheet.Cells(row + 2, 21) = obj.SECC_Other_4
            Next

            ' 자동 너비 조정
            worksheet.Columns.AutoFit()

            ' 헤더 스타일
            Dim headerRange = worksheet.Range("A1", worksheet.Cells(1, headers.Length))
            headerRange.Font.Bold = True
            headerRange.Interior.Color = RGB(200, 200, 200)

            workbook.SaveAs(filePath)
            Logger.Log($"엑셀 내보내기 완료: {data.Count}개 행")

        Catch ex As Exception
            Logger.LogError($"엑셀 내보내기 오류: {ex.Message}")
            Throw
        Finally
            If workbook IsNot Nothing Then
                workbook.Close(False)
                Marshal.ReleaseComObject(workbook)
            End If
            If excelApp IsNot Nothing Then
                excelApp.Quit()
                Marshal.ReleaseComObject(excelApp)
            End If
        End Try
    End Sub

    Public Shared Sub ExportMappingResults(results As List(Of MatchResult), filePath As String)
        Dim excelApp As Application = Nothing
        Dim workbook As Workbook = Nothing
        Dim worksheet As Worksheet = Nothing

        Try
            Logger.Log($"매핑 결과 엑셀 내보내기 시작: {filePath}")

            excelApp = New Application()
            excelApp.Visible = False
            excelApp.DisplayAlerts = False

            workbook = excelApp.Workbooks.Add()
            worksheet = CType(workbook.Worksheets(1), Worksheet)
            worksheet.Name = "Mapping Results"

            ' 헤더 작성
            Dim headers = {"StructuralLinkName", "GenericLinkName", "StructuralElementId", "GenericElementId",
                          "StructuralX", "StructuralY", "StructuralZ", "GenericX", "GenericY", "GenericZ",
                          "Distance(mm)", "ColumnNumber", "IsPerfectMatch", "Comment"}

            For col As Integer = 0 To headers.Length - 1
                worksheet.Cells(1, col + 1) = headers(col)
            Next

            ' 데이터 작성
            For row As Integer = 0 To results.Count - 1
                Dim res = results(row)
                worksheet.Cells(row + 2, 1) = res.StructuralLinkName
                worksheet.Cells(row + 2, 2) = res.GenericLinkName
                worksheet.Cells(row + 2, 3) = res.StructuralElementId
                worksheet.Cells(row + 2, 4) = res.GenericElementId
                worksheet.Cells(row + 2, 5) = res.StructuralX
                worksheet.Cells(row + 2, 6) = res.StructuralY
                worksheet.Cells(row + 2, 7) = res.StructuralZ
                worksheet.Cells(row + 2, 8) = res.GenericX
                worksheet.Cells(row + 2, 9) = res.GenericY
                worksheet.Cells(row + 2, 10) = res.GenericZ
                worksheet.Cells(row + 2, 11) = res.Distance
                worksheet.Cells(row + 2, 12) = res.ColumnNumber
                worksheet.Cells(row + 2, 13) = If(res.IsPerfectMatch, "Yes", "No")
                worksheet.Cells(row + 2, 14) = res.Comment

                ' 검토 필요 항목 하이라이트
                If Not res.IsPerfectMatch Then
                    Dim rowRange = worksheet.Range(worksheet.Cells(row + 2, 1), worksheet.Cells(row + 2, headers.Length))
                    rowRange.Interior.Color = RGB(255, 255, 200) ' 연한 노란색
                End If
            Next

            ' 자동 너비 조정
            worksheet.Columns.AutoFit()

            ' 헤더 스타일
            Dim headerRange = worksheet.Range("A1", worksheet.Cells(1, headers.Length))
            headerRange.Font.Bold = True
            headerRange.Interior.Color = RGB(200, 200, 200)

            workbook.SaveAs(filePath)
            Logger.Log($"매핑 결과 엑셀 내보내기 완료: {results.Count}개 행")

        Catch ex As Exception
            Logger.LogError($"매핑 결과 엑셀 내보내기 오류: {ex.Message}")
            Throw
        Finally
            If workbook IsNot Nothing Then
                workbook.Close(False)
                Marshal.ReleaseComObject(workbook)
            End If
            If excelApp IsNot Nothing Then
                excelApp.Quit()
                Marshal.ReleaseComObject(excelApp)
            End If
        End Try
    End Sub

    Public Shared Function ImportFromExcel(filePath As String) As List(Of ObjectData)
        Dim results As New List(Of ObjectData)
        Dim excelApp As Application = Nothing
        Dim workbook As Workbook = Nothing
        Dim worksheet As Worksheet = Nothing

        Try
            Logger.Log($"엑셀 임포트 시작: {filePath}")

            excelApp = New Application()
            excelApp.Visible = False
            excelApp.DisplayAlerts = False

            workbook = excelApp.Workbooks.Open(filePath)
            worksheet = CType(workbook.Worksheets(1), Worksheet)

            Dim lastRow = worksheet.Cells(worksheet.Rows.Count, 1).End(XlDirection.xlUp).Row

            For row As Integer = 2 To lastRow
                Dim obj As New ObjectData With {
                    .LinkFileName = GetCellValue(worksheet.Cells(row, 1)),
                    .ElementId = GetCellValue(worksheet.Cells(row, 2)),
                    .ElementName = GetCellValue(worksheet.Cells(row, 3)),
                    .CategoryName = GetCellValue(worksheet.Cells(row, 4)),
                    .LocationX = GetDoubleValue(worksheet.Cells(row, 5)),
                    .LocationY = GetDoubleValue(worksheet.Cells(row, 6)),
                    .LocationZ = GetDoubleValue(worksheet.Cells(row, 7)),
                    .Grid = GetCellValue(worksheet.Cells(row, 8)),
                    .SB_NAME = GetCellValue(worksheet.Cells(row, 9)),
                    .SB_FIELD = GetCellValue(worksheet.Cells(row, 10)),
                    .SB_FL = GetCellValue(worksheet.Cells(row, 11)),
                    .SB_BLDG = GetCellValue(worksheet.Cells(row, 12)),
                    .StudNumber = GetCellValue(worksheet.Cells(row, 13)),
                    .ColumnNumber = GetCellValue(worksheet.Cells(row, 14)),
                    .EXCLUSION = GetCellValue(worksheet.Cells(row, 15)),
                    .CLASS = GetCellValue(worksheet.Cells(row, 16)),
                    .SAMOO_Other_1 = GetCellValue(worksheet.Cells(row, 17)),
                    .SAMOO_Other_2 = GetCellValue(worksheet.Cells(row, 18)),
                    .SAMOO_Other_3 = GetCellValue(worksheet.Cells(row, 19)),
                    .SECC_Other_3 = GetCellValue(worksheet.Cells(row, 20)),
                    .SECC_Other_4 = GetCellValue(worksheet.Cells(row, 21))
                }
                results.Add(obj)
            Next

            Logger.Log($"엑셀 임포트 완료: {results.Count}개 행")

        Catch ex As Exception
            Logger.LogError($"엑셀 임포트 오류: {ex.Message}")
            Throw
        Finally
            If workbook IsNot Nothing Then
                workbook.Close(False)
                Marshal.ReleaseComObject(workbook)
            End If
            If excelApp IsNot Nothing Then
                excelApp.Quit()
                Marshal.ReleaseComObject(excelApp)
            End If
        End Try

        Return results
    End Function

    Public Shared Function ImportMappingResults(filePath As String) As List(Of MatchResult)
        Dim results As New List(Of MatchResult)
        Dim excelApp As Application = Nothing
        Dim workbook As Workbook = Nothing
        Dim worksheet As Worksheet = Nothing

        Try
            Logger.Log($"매핑 결과 엑셀 임포트 시작: {filePath}")

            excelApp = New Application()
            excelApp.Visible = False
            excelApp.DisplayAlerts = False

            workbook = excelApp.Workbooks.Open(filePath)
            worksheet = CType(workbook.Worksheets(1), Worksheet)

            Dim lastRow = worksheet.Cells(worksheet.Rows.Count, 1).End(XlDirection.xlUp).Row

            For row As Integer = 2 To lastRow
                Dim res As New MatchResult With {
                    .StructuralLinkName = GetCellValue(worksheet.Cells(row, 1)),
                    .GenericLinkName = GetCellValue(worksheet.Cells(row, 2)),
                    .StructuralElementId = GetCellValue(worksheet.Cells(row, 3)),
                    .GenericElementId = GetCellValue(worksheet.Cells(row, 4)),
                    .ColumnNumber = GetCellValue(worksheet.Cells(row, 12))
                }
                results.Add(res)
            Next

            Logger.Log($"매핑 결과 엑셀 임포트 완료: {results.Count}개 행")

        Catch ex As Exception
            Logger.LogError($"매핑 결과 엑셀 임포트 오류: {ex.Message}")
            Throw
        Finally
            If workbook IsNot Nothing Then
                workbook.Close(False)
                Marshal.ReleaseComObject(workbook)
            End If
            If excelApp IsNot Nothing Then
                excelApp.Quit()
                Marshal.ReleaseComObject(excelApp)
            End If
        End Try

        Return results
    End Function

    Private Shared Function GetCellValue(cell As Range) As String
        If cell.Value IsNot Nothing Then
            Return cell.Value.ToString()
        End If
        Return ""
    End Function

    Private Shared Function GetDoubleValue(cell As Range) As Double
        If cell.Value IsNot Nothing Then
            Dim value As Double
            If Double.TryParse(cell.Value.ToString(), value) Then
                Return value
            End If
        End If
        Return 0
    End Function
End Class
