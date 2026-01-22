Imports Autodesk.Revit.DB
Imports System.Collections.Generic
Imports System.Linq

Public Class ObjectMatcher
    Private tolerance As Double ' feet 단위

    Public Sub New(toleranceInFeet As Double)
        tolerance = toleranceInFeet
    End Sub

    Public Function MatchObjects(structuralData As List(Of ObjectData),
                                  genericData As List(Of ObjectData),
                                  structuralLinks As List(Of RevitLinkInstance),
                                  genericLinks As List(Of RevitLinkInstance)) As List(Of MatchResult)

        Dim results As New List(Of MatchResult)

        Logger.Log("=== 매칭 시작 ===")
        Logger.Log($"Structural Columns: {structuralData.Count}개")
        Logger.Log($"Generic Models: {genericData.Count}개")

        ' 링크 파일 쌍 찾기
        Dim linkPairs = FindLinkPairs(structuralLinks, genericLinks)

        Logger.Log($"발견된 링크 파일 쌍: {linkPairs.Count}개")

        For Each pair In linkPairs
            Logger.Log($"매칭 중: {pair.StructuralLinkName} <-> {pair.GenericLinkName}")

            ' 해당 링크에서 데이터 필터링
            Dim structObjs = structuralData.Where(Function(d) d.LinkFileName = pair.StructuralLinkName).ToList()
            Dim genObjs = genericData.Where(Function(d) d.LinkFileName = pair.GenericLinkName).ToList()

            Logger.Log($"  Structural: {structObjs.Count}개, Generic: {genObjs.Count}개")

            ' XY 좌표 기반 매칭
            For Each structObj In structObjs
                Dim match = FindBestMatch(structObj, genObjs)

                If match IsNot Nothing Then
                    Dim matchResult As New MatchResult With {
                        .StructuralLinkName = pair.StructuralLinkName,
                        .GenericLinkName = pair.GenericLinkName,
                        .StructuralElementId = structObj.ElementId,
                        .GenericElementId = match.GenericObj.ElementId,
                        .StructuralX = structObj.LocationX,
                        .StructuralY = structObj.LocationY,
                        .StructuralZ = structObj.LocationZ,
                        .GenericX = match.GenericObj.LocationX,
                        .GenericY = match.GenericObj.LocationY,
                        .GenericZ = match.GenericObj.LocationZ,
                        .Distance = match.Distance,
                        .ColumnNumber = match.GenericObj.ColumnNumber,
                        .IsPerfectMatch = match.IsPerfect,
                        .Comment = If(match.IsPerfect, "완벽 일치", "검토 필요 - 좌표 차이 있음")
                    }

                    results.Add(matchResult)
                End If
            Next
        Next

        Dim perfectCount = results.Count(Function(r) r.IsPerfectMatch)
        Dim reviewCount = results.Count(Function(r) Not r.IsPerfectMatch)

        Logger.Log($"=== 매칭 완료 ===")
        Logger.Log($"완벽 일치: {perfectCount}개")
        Logger.Log($"검토 필요: {reviewCount}개")
        Logger.Log($"총 매칭: {results.Count}개")

        Return results
    End Function

    Private Function FindLinkPairs(structuralLinks As List(Of RevitLinkInstance),
                                    genericLinks As List(Of RevitLinkInstance)) As List(Of LinkPair)

        Dim pairs As New List(Of LinkPair)

        For Each sLink In structuralLinks
            Dim sName = sLink.Name

            ' 파일명에서 키워드 추출 (예: 1_A_a.rvt -> A)
            Dim sKeyword = ExtractKeyword(sName, "S")
            If String.IsNullOrEmpty(sKeyword) Then Continue For

            ' 매칭되는 Generic 링크 찾기
            For Each gLink In genericLinks
                Dim gName = gLink.Name
                Dim gKeyword = ExtractKeyword(gName, "A")

                If String.IsNullOrEmpty(gKeyword) Then Continue For

                ' 키워드만 다르고 나머지가 같은지 확인
                If AreMatchingLinks(sName, gName, sKeyword, gKeyword) Then
                    pairs.Add(New LinkPair With {
                        .StructuralLinkName = sName,
                        .GenericLinkName = gName
                    })
                    Logger.Log($"링크 쌍 발견: {sName} <-> {gName}")
                    Exit For
                End If
            Next
        Next

        Return pairs
    End Function

    Private Function ExtractKeyword(fileName As String, expectedKeyword As String) As String
        ' 파일명 패턴: ***_XXX_*** 형태에서 가운데 키워드 추출
        Dim parts = fileName.Split("_"c)

        For i = 0 To parts.Length - 1
            If parts(i).StartsWith(expectedKeyword, StringComparison.OrdinalIgnoreCase) Then
                Return parts(i)
            End If
        Next

        ' 정확히 일치하는 부분 찾기
        For Each Part In parts
            If Part.Equals(expectedKeyword, StringComparison.OrdinalIgnoreCase) Then
                Return Part
            End If
        Next

        Return ""
    End Function

    Private Function AreMatchingLinks(sName As String, gName As String, sKeyword As String, gKeyword As String) As Boolean
        ' 키워드만 다르고 나머지가 동일한지 확인
        Dim sPattern = sName.Replace(sKeyword, "{KEYWORD}")
        Dim gPattern = gName.Replace(gKeyword, "{KEYWORD}")

        Return sPattern.Equals(gPattern, StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function FindBestMatch(structObj As ObjectData, genObjs As List(Of ObjectData)) As MatchInfo
        Dim bestMatch As MatchInfo = Nothing
        Dim minDistance As Double = Double.MaxValue

        For Each genObj In genObjs
            ' XY 평면 거리 계산 (feet 단위)
            Dim dx = (structObj.LocationX - genObj.LocationX) / 304.8 ' mm to feet
            Dim dy = (structObj.LocationY - genObj.LocationY) / 304.8
            Dim distance = Math.Sqrt(dx * dx + dy * dy)

            ' 허용 오차 내에서 가장 가까운 객체 찾기
            If distance <= tolerance AndAlso distance < minDistance Then
                minDistance = distance
                bestMatch = New MatchInfo With {
                    .GenericObj = genObj,
                    .Distance = Math.Round(distance * 304.8, 2), ' feet to mm
                    .IsPerfect = (distance < 0.001) ' 약 0.3mm 이내면 완벽 일치
                }
            End If
        Next

        Return bestMatch
    End Function
End Class

Public Class LinkPair
    Public Property StructuralLinkName As String
    Public Property GenericLinkName As String
End Class

Public Class MatchInfo
    Public Property GenericObj As ObjectData
    Public Property Distance As Double ' mm 단위
    Public Property IsPerfect As Boolean
End Class

Public Class MatchResult
    Public Property StructuralLinkName As String = ""
    Public Property GenericLinkName As String = ""
    Public Property StructuralElementId As String = ""
    Public Property GenericElementId As String = ""
    Public Property StructuralX As Double = 0
    Public Property StructuralY As Double = 0
    Public Property StructuralZ As Double = 0
    Public Property GenericX As Double = 0
    Public Property GenericY As Double = 0
    Public Property GenericZ As Double = 0
    Public Property Distance As Double = 0 ' mm 단위
    Public Property ColumnNumber As String = ""
    Public Property IsPerfectMatch As Boolean = False
    Public Property Comment As String = ""
End Class