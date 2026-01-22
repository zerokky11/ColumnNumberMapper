Imports System.Drawing
Imports System.Windows.Forms
Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI

<Transaction(TransactionMode.Manual)>
<Regeneration(RegenerationOption.Manual)>
Public Class ColumnNumberMapper
    Implements IExternalCommand

    Public Function Execute(commandData As ExternalCommandData, ByRef message As String, elements As ElementSet) As Result Implements IExternalCommand.Execute
        Dim uiApp As UIApplication = commandData.Application
        Dim uiDoc As UIDocument = uiApp.ActiveUIDocument
        Dim doc As Document = uiDoc.Document

        Logger.Initialize()
        Logger.Log("=== Column Number Mapper 시작 ===")

        Try
            ' 메인 폼 표시
            Using mainForm As New MainForm(uiDoc)
                If mainForm.ShowDialog() = DialogResult.OK Then
                    Logger.Log("프로세스 완료")
                    TaskDialog.Show("완료", "작업이 성공적으로 완료되었습니다.")
                    Return Result.Succeeded
                Else
                    Logger.Log("사용자가 취소함")
                    Return Result.Cancelled
                End If
            End Using

        Catch ex As Exception
            Logger.LogError("Execute 오류: " & ex.Message)
            TaskDialog.Show("오류", "오류가 발생했습니다: " & ex.Message)
            Return Result.Failed
        End Try
    End Function
End Class

Public Class MainForm
    Inherits Form

    Private uiDoc As UIDocument
    Private doc As Document
    Private btnStep1 As Button
    Private btnStep2 As Button
    Private btnStep3 As Button
    Private btnStep4 As Button
    Private lblStatus As Label
    Private txtTolerance As TextBox
    Private lblTolerance As Label

    Private structuralLinksData As New List(Of RevitLinkInstance)
    Private genericLinksData As New List(Of RevitLinkInstance)
    Private step1ExcelPath As String = ""
    Private step2ExcelPath As String = ""
    Private step3ExcelPath As String = ""

    Public Sub New(ByVal uiDocument As UIDocument)
        uiDoc = uiDocument
        doc = uiDoc.Document
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Column Number Mapper"
        Me.Size = New Size(500, 400)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        Dim y As Integer = 20

        ' 허용 오차 입력
        lblTolerance = New Label With {
            .Text = "XY 좌표 허용 오차 (mm):",
            .Location = New Point(20, y),
            .Size = New Size(150, 20)
        }
        Me.Controls.Add(lblTolerance)

        txtTolerance = New TextBox With {
            .Text = "50",
            .Location = New Point(180, y),
            .Size = New Size(80, 20)
        }
        Me.Controls.Add(txtTolerance)
        y += 40

        ' Step 1 버튼
        btnStep1 = New Button With {
            .Text = "Step 1: Structural Columns 추출",
            .Location = New Point(20, y),
            .Size = New Size(450, 40)
        }
        AddHandler btnStep1.Click, AddressOf BtnStep1_Click
        Me.Controls.Add(btnStep1)
        y += 50

        ' Step 2 버튼
        btnStep2 = New Button With {
            .Text = "Step 2: Generic Models 추출",
            .Location = New Point(20, y),
            .Size = New Size(450, 40),
            .Enabled = False
        }
        AddHandler btnStep2.Click, AddressOf BtnStep2_Click
        Me.Controls.Add(btnStep2)
        y += 50

        ' Step 3 버튼
        btnStep3 = New Button With {
            .Text = "Step 3: 매핑 분석 및 결과 출력",
            .Location = New Point(20, y),
            .Size = New Size(450, 40),
            .Enabled = False
        }
        AddHandler btnStep3.Click, AddressOf BtnStep3_Click
        Me.Controls.Add(btnStep3)
        y += 50

        ' Step 4 버튼
        btnStep4 = New Button With {
            .Text = "Step 4: S5_EQCODE 값 업데이트",
            .Location = New Point(20, y),
            .Size = New Size(450, 40),
            .Enabled = False
        }
        AddHandler btnStep4.Click, AddressOf BtnStep4_Click
        Me.Controls.Add(btnStep4)
        y += 50

        ' 상태 표시 레이블
        lblStatus = New Label With {
            .Text = "Step 1부터 시작하세요.",
            .Location = New Point(20, y),
            .Size = New Size(450, 60),
            .BorderStyle = BorderStyle.FixedSingle,
            .TextAlign = CONTENTALIGNMENT.MiddleLeft
        }
        Me.Controls.Add(lblStatus)
    End Sub

    Private Sub BtnStep1_Click(sender As Object, e As EventArgs)
        Logger.Log("--- Step 1 시작: Structural Columns 추출 ---")

        Try
            ' 링크 파일 선택
            Dim selector As New LinkSelector(uiDoc)
            structuralLinksData = selector.SelectLinks("Structural Columns를 추출할 링크 파일을 선택하세요.")

            If structuralLinksData.Count = 0 Then
                MessageBox.Show("선택된 링크가 없습니다.")
                Return
            End If

            Logger.Log($"선택된 링크 수: {structuralLinksData.Count}")

            ' 데이터 추출
            Dim extractor As New DataExtractor()
            Dim data = extractor.ExtractData(structuralLinksData, BuiltInCategory.OST_StructuralColumns, doc)

            ' 엑셀로 내보내기
            Dim sfd As New SaveFileDialog With {
                .Filter = "Excel Files|*.xlsx",
                .FileName = "Step1_StructuralColumns_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx"
            }

            If sfd.ShowDialog() = DialogResult.OK Then
                step1ExcelPath = sfd.FileName
                ExcelHelper.ExportToExcel(data, step1ExcelPath, "Structural Columns")
                lblStatus.Text = $"Step 1 완료! {data.Count}개 객체 추출됨." & vbCrLf & $"파일: {step1ExcelPath}"
                btnStep2.Enabled = True
                Logger.Log($"Step 1 완료: {data.Count}개 객체, 파일: {step1ExcelPath}")
            End If

        Catch ex As Exception
            Logger.LogError("Step 1 오류: " & ex.Message)
            MessageBox.Show("오류: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnStep2_Click(sender As Object, e As EventArgs)
        Logger.Log("--- Step 2 시작: Generic Models 추출 ---")

        Try
            ' 링크 파일 선택
            Dim selector As New LinkSelector(uiDoc)
            genericLinksData = selector.SelectLinks("Generic Models를 추출할 링크 파일을 선택하세요.")

            If genericLinksData.Count = 0 Then
                MessageBox.Show("선택된 링크가 없습니다.")
                Return
            End If

            Logger.Log($"선택된 링크 수: {genericLinksData.Count}")

            ' 데이터 추출
            Dim extractor As New DataExtractor()
            Dim data = extractor.ExtractData(genericLinksData, BuiltInCategory.OST_GenericModel, doc)

            ' 엑셀로 내보내기
            Dim sfd As New SaveFileDialog With {
                .Filter = "Excel Files|*.xlsx",
                .FileName = "Step2_GenericModels_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx"
            }

            If sfd.ShowDialog() = DialogResult.OK Then
                step2ExcelPath = sfd.FileName
                ExcelHelper.ExportToExcel(data, step2ExcelPath, "Generic Models")
                lblStatus.Text = $"Step 2 완료! {data.Count}개 객체 추출됨." & vbCrLf & $"파일: {step2ExcelPath}"
                btnStep3.Enabled = True
                Logger.Log($"Step 2 완료: {data.Count}개 객체, 파일: {step2ExcelPath}")
            End If

        Catch ex As Exception
            Logger.LogError("Step 2 오류: " & ex.Message)
            MessageBox.Show("오류: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnStep3_Click(sender As Object, e As EventArgs)
        Logger.Log("--- Step 3 시작: 매핑 분석 ---")

        Try
            Dim tolerance As Double
            If Not Double.TryParse(txtTolerance.Text, tolerance) Then
                MessageBox.Show("유효한 허용 오차를 입력하세요.")
                Return
            End If

            ' 허용 오차를 feet 단위로 변환 (Revit 내부 단위)
            Dim toleranceFeet As Double = tolerance / 304.8

            Logger.Log($"허용 오차: {tolerance}mm ({toleranceFeet}ft)")

            ' 데이터 로드
            Dim structuralData = ExcelHelper.ImportFromExcel(step1ExcelPath)
            Dim genericData = ExcelHelper.ImportFromExcel(step2ExcelPath)

            ' 매칭 수행
            Dim matcher As New ObjectMatcher(toleranceFeet)
            Dim matchResults = matcher.MatchObjects(structuralData, genericData, structuralLinksData, genericLinksData)

            ' 결과를 엑셀로 내보내기
            Dim sfd As New SaveFileDialog With {
                .Filter = "Excel Files|*.xlsx",
                .FileName = "Step3_MappingResults_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx"
            }

            If sfd.ShowDialog() = DialogResult.OK Then
                step3ExcelPath = sfd.FileName
                ExcelHelper.ExportMappingResults(matchResults, step3ExcelPath)

                Dim perfectCount = matchResults.Count(Function(m) m.IsPerfectMatch)
                Dim needReviewCount = matchResults.Count(Function(m) Not m.IsPerfectMatch)

                lblStatus.Text = $"Step 3 완료!" & vbCrLf &
                                $"완벽 매칭: {perfectCount}개, 검토 필요: {needReviewCount}개" & vbCrLf &
                                $"파일: {step3ExcelPath}" & vbCrLf &
                                "검토 후 불필요한 행을 삭제하고 Step 4로 진행하세요."
                btnStep4.Enabled = True

                Logger.Log($"Step 3 완료: 완벽 매칭 {perfectCount}, 검토 필요 {needReviewCount}")
            End If

        Catch ex As Exception
            Logger.LogError("Step 3 오류: " & ex.Message)
            MessageBox.Show("오류: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnStep4_Click(sender As Object, e As EventArgs)
        Logger.Log("--- Step 4 시작: S5_EQCODE 업데이트 ---")

        Try
            ' 검토 완료된 엑셀 파일 선택
            Dim ofd As New OpenFileDialog With {
                .Filter = "Excel Files|*.xlsx",
                .Title = "검토 완료된 매핑 결과 파일을 선택하세요"
            }

            If ofd.ShowDialog() <> DialogResult.OK Then
                Return
            End If

            ' 매핑 데이터 로드
            Dim mappingData = ExcelHelper.ImportMappingResults(ofd.FileName)

            Logger.Log($"로드된 매핑 데이터: {mappingData.Count}개")

            ' 업데이트할 링크 파일들 선택
            Dim selector As New LinkSelector(uiDoc)
            Dim selectedLinks = selector.SelectLinks("S5_EQCODE를 업데이트할 링크 파일들을 선택하세요.")

            If selectedLinks.Count = 0 Then
                MessageBox.Show("선택된 링크가 없습니다.")
                Return
            End If

            ' 파라미터 업데이트 수행
            Dim updater As New ParameterUpdater()
            Dim result = updater.UpdateParameters(selectedLinks, mappingData, doc)

            lblStatus.Text = $"Step 4 완료!" & vbCrLf &
                            $"업데이트 성공: {result.SuccessCount}개" & vbCrLf &
                            $"실패: {result.FailCount}개" & vbCrLf &
                            $"자세한 내용은 로그를 확인하세요."

            Logger.Log($"Step 4 완료: 성공 {result.SuccessCount}, 실패 {result.FailCount}")

            MessageBox.Show($"업데이트 완료!" & vbCrLf &
                           $"성공: {result.SuccessCount}개" & vbCrLf &
                           $"실패: {result.FailCount}개", "완료")

            Me.DialogResult = DialogResult.OK
            Me.Close()

        Catch ex As Exception
            Logger.LogError("Step 4 오류: " & ex.Message)
            MessageBox.Show("오류: " & ex.Message)
        End Try
    End Sub
End Class