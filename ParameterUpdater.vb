Imports System.Collections.Generic
Imports System.Linq
Imports System.Windows.Forms.LinkLabel
Imports Autodesk.Revit.DB

Public Class ParameterUpdater

    Public Function UpdateParameters(links As List(Of RevitLinkInstance),
                                      mappingData As List(Of MatchResult),
                                      hostDoc As Document) As UpdateResult

        Dim result As New UpdateResult()

        Logger.Log("=== S5_EQCODE 업데이트 시작 ===")
        Logger.Log($"대상 링크: {links.Count}개")
        Logger.Log($"매핑 데이터: {mappingData.Count}개")

        For Each link In links
            Logger.Log($"링크 파일 처리 중: {link.Name}")

            Dim linkDoc As Document = link.GetLinkDocument()
            If linkDoc Is Nothing Then
                Logger.LogError($"링크 문서를 열 수 없음: {link.Name}")
                Continue For
            End If

            ' 이 링크에 해당하는 매핑 데이터 필터링
            Dim relevantMappings = mappingData.Where(Function(m) m.StructuralLinkName = link.Name).ToList()

            If relevantMappings.Count = 0 Then
                Logger.Log($"  - 매핑 데이터 없음, 건너뜀")
                Continue For
            End If

            Logger.Log($"  - {relevantMappings.Count}개 매핑 발견")

            ' Transaction 시작
            Using trans As New Transaction(linkDoc, "Update S5_EQCODE")
                Try
                    trans.Start()

                    For Each mapping In relevantMappings
                        Try
                            ' Element ID로 요소 찾기
                            Dim elemId As New ElementId(Integer.Parse(mapping.StructuralElementId))
                            Dim elem = linkDoc.GetElement(elemId)

                            If elem Is Nothing Then
                                Logger.LogError($"    요소를 찾을 수 없음: ID {mapping.StructuralElementId}")
                                result.FailCount += 1
                                Continue For
                            End If

                            ' S5_EQCODE 파라미터 찾기
                            Dim param = elem.LookupParameter("S5_EQCODE")

                            If param Is Nothing Then
                                Logger.LogError($"    S5_EQCODE 파라미터 없음: ID {mapping.StructuralElementId}")
                                result.FailCount += 1
                                Continue For
                            End If

                            If param.IsReadOnly Then
                                Logger.LogError($"    S5_EQCODE 읽기 전용: ID {mapping.StructuralElementId}")
                                result.FailCount += 1
                                Continue For
                            End If

                            ' 값 설정
                            param.Set(mapping.ColumnNumber)

                            Logger.Log($"    ✓ ID {mapping.StructuralElementId}: S5_EQCODE = '{mapping.ColumnNumber}'")
                            result.SuccessCount += 1

                        Catch ex As Exception
                            Logger.LogError($"    요소 업데이트 오류 (ID: {mapping.StructuralElementId}): {ex.Message}")
                            result.FailCount += 1
                        End Try
                    Next

                    trans.Commit()
                    Logger.Log($"  - Transaction 커밋 완료")

                Catch ex As Exception
                    trans.RollBack()
                    Logger.LogError($"  - Transaction 롤백: {ex.Message}")
                End Try
            End Using
        Next

        Logger.Log("=== 업데이트 완료 ===")
        Logger.Log($"성공: {result.SuccessCount}개")
        Logger.Log($"실패: {result.FailCount}개")

        Return result
    End Function
End Class

Public Class UpdateResult
    Public Property SuccessCount As Integer = 0
    Public Property FailCount As Integer = 0
End Class