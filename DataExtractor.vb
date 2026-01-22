Imports System.Collections.Generic
Imports Autodesk.Revit.DB

Public Class DataExtractor

    Public Function ExtractData(links As List(Of RevitLinkInstance), category As BuiltInCategory, hostDoc As Document) As List(Of ObjectData)
        Dim results As New List(Of ObjectData)

        For Each link As RevitLinkInstance In links
            Logger.Log($"링크 파일 처리 중: {link.Name}")

            Dim linkDoc As Document = link.GetLinkDocument()
            If linkDoc Is Nothing Then
                Logger.LogError($"링크 문서를 열 수 없음: {link.Name}")
                Continue For
            End If

            ' 카테고리 필터
            Dim collector As New FilteredElementCollector(linkDoc)
            Dim elements As IList(Of Element) = collector.OfCategory(category).WhereElementIsNotElementType().ToElements()

            Logger.Log($"  - {elements.Count}개 요소 발견")

            For Each elem As Element In elements
                Try
                    Dim objData As New ObjectData()
                    objData.LinkFileName = link.Name
                    objData.ElementId = elem.Id.IntegerValue.ToString()

                    ' Type Name
                    Dim typeId = elem.GetTypeId()
                    If typeId <> ElementId.InvalidElementId Then
                        Dim elemType = linkDoc.GetElement(typeId)
                        If elemType IsNot Nothing Then
                            objData.ElementName = elemType.Name
                        End If
                    End If

                    ' Category Name
                    If elem.Category IsNot Nothing Then
                        objData.CategoryName = elem.Category.Name
                    End If

                    ' Location (XYZ)
                    Dim location = GetElementLocation(elem)
                    If location IsNot Nothing Then
                        ' 링크 Transform 적용
                        Dim transform = link.GetTotalTransform()
                        Dim transformedLoc = transform.OfPoint(location)

                        ' mm 단위로 변환 (Revit 내부는 feet)
                        objData.LocationX = Math.Round(transformedLoc.X * 304.8, 2)
                        objData.LocationY = Math.Round(transformedLoc.Y * 304.8, 2)
                        objData.LocationZ = Math.Round(transformedLoc.Z * 304.8, 2)
                    End If

                    ' 파라미터 추출
                    objData.Grid = GetParameterValue(elem, "Grid")
                    objData.SB_NAME = GetParameterValue(elem, "SB_NAME")
                    objData.SB_FIELD = GetParameterValue(elem, "SB_FIELD")
                    objData.SB_FL = GetParameterValue(elem, "SB_FL")
                    objData.SB_BLDG = GetParameterValue(elem, "SB_BLDG")
                    objData.StudNumber = GetParameterValue(elem, "Stud Number")
                    objData.ColumnNumber = GetParameterValue(elem, "Column Number")
                    objData.EXCLUSION = GetParameterValue(elem, "EXCLUSION")
                    objData.CLASS = GetParameterValue(elem, "CLASS")
                    objData.SAMOO_Other_1 = GetParameterValue(elem, "SAMOO-Other_1")
                    objData.SAMOO_Other_2 = GetParameterValue(elem, "SAMOO-Other_2")
                    objData.SAMOO_Other_3 = GetParameterValue(elem, "SAMOO-Other_3")
                    objData.SECC_Other_3 = GetParameterValue(elem, "SECC-Other_3")
                    objData.SECC_Other_4 = GetParameterValue(elem, "SECC-Other_4")

                    results.Add(objData)

                Catch ex As Exception
                    Logger.LogError($"  - 요소 처리 오류 (ID: {elem.Id}): {ex.Message}")
                End Try
            Next
        Next

        Logger.Log($"총 {results.Count}개 객체 데이터 추출 완료")
        Return results
    End Function

    Private Function GetElementLocation(elem As Element) As XYZ
        ' LocationPoint 시도
        Dim locPoint = TryCast(elem.Location, LocationPoint)
        If locPoint IsNot Nothing Then
            Return locPoint.Point
        End If

        ' LocationCurve 시도 (중점 사용)
        Dim locCurve = TryCast(elem.Location, LocationCurve)
        If locCurve IsNot Nothing Then
            Dim curve = locCurve.Curve
            Return curve.Evaluate(0.5, True)
        End If

        ' BoundingBox 중심점 시도
        Dim bb = elem.get_BoundingBox(Nothing)
        If bb IsNot Nothing Then
            Return (bb.Min + bb.Max) / 2
        End If

        Return Nothing
    End Function

    Private Function GetParameterValue(elem As Element, paramName As String) As String
        Try
            Dim param = elem.LookupParameter(paramName)
            If param IsNot Nothing AndAlso param.HasValue Then
                Select Case param.StorageType
                    Case StorageType.String
                        Return param.AsString()
                    Case StorageType.Integer
                        Return param.AsInteger().ToString()
                    Case StorageType.Double
                        Return param.AsDouble().ToString()
                    Case StorageType.ElementId
                        Return param.AsElementId().IntegerValue.ToString()
                End Select
            End If
        Catch ex As Exception
            ' 파라미터가 없거나 오류 발생 시 빈 문자열 반환
        End Try
        Return ""
    End Function
End Class

Public Class ObjectData
    Public Property LinkFileName As String = ""
    Public Property ElementId As String = ""
    Public Property ElementName As String = ""
    Public Property CategoryName As String = ""
    Public Property LocationX As Double = 0
    Public Property LocationY As Double = 0
    Public Property LocationZ As Double = 0
    Public Property Grid As String = ""
    Public Property SB_NAME As String = ""
    Public Property SB_FIELD As String = ""
    Public Property SB_FL As String = ""
    Public Property SB_BLDG As String = ""
    Public Property StudNumber As String = ""
    Public Property ColumnNumber As String = ""
    Public Property EXCLUSION As String = ""
    Public Property [CLASS] As String = ""
    Public Property SAMOO_Other_1 As String = ""
    Public Property SAMOO_Other_2 As String = ""
    Public Property SAMOO_Other_3 As String = ""
    Public Property SECC_Other_3 As String = ""
    Public Property SECC_Other_4 As String = ""
End Class
