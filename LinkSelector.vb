Imports System.Collections.Generic
Imports System.Drawing
Imports System.Linq
Imports System.Windows.Forms
Imports System.Windows.Forms.LinkLabel
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI

Public Class LinkSelector
    Private uiDoc As UIDocument
    Private doc As Document

    Public Sub New(ByVal uiDocument As UIDocument)
        uiDoc = uiDocument
        doc = uiDoc.Document
    End Sub

    Public Function SelectLinks(prompt As String) As List(Of RevitLinkInstance)
        Dim selectedLinks As New List(Of RevitLinkInstance)

        ' 프로젝트의 모든 링크 수집
        Dim collector As New FilteredElementCollector(doc)
        Dim allLinks = collector.OfClass(GetType(RevitLinkInstance)).Cast(Of RevitLinkInstance).ToList()

        If allLinks.Count = 0 Then
            MessageBox.Show("프로젝트에 링크 파일이 없습니다.")
            Logger.Log("링크 파일 없음")
            Return selectedLinks
        End If

        Logger.Log($"발견된 링크 파일: {allLinks.Count}개")

        ' 링크 선택 다이얼로그 표시
        Using selectForm As New LinkSelectionForm(allLinks, prompt)
            If selectForm.ShowDialog() = DialogResult.OK Then
                selectedLinks = selectForm.SelectedLinks
                Logger.Log($"사용자가 선택한 링크: {selectedLinks.Count}개")

                For Each link In selectedLinks
                    Logger.Log($"  - {link.Name}")
                Next
            Else
                Logger.Log("링크 선택 취소됨")
            End If
        End Using

        Return selectedLinks
    End Function
End Class

Public Class LinkSelectionForm
    Inherits System.Windows.Forms.Form

    Private allLinks As List(Of RevitLinkInstance)
    Private checkedListBox As CheckedListBox
    Private btnOK As Button
    Private btnCancel As Button
    Private btnSelectAll As Button
    Private btnDeselectAll As Button

    Public Property SelectedLinks As New List(Of RevitLinkInstance)

    Public Sub New(links As List(Of RevitLinkInstance), prompt As String)
        allLinks = links
        InitializeComponent(prompt)
    End Sub

    Private Sub InitializeComponent(prompt As String)
        Me.Text = "링크 파일 선택"
        Me.Size = New Size(600, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False

        Dim y As Integer = 10

        ' 안내 레이블
        Dim lblPrompt As New Label With {
            .Text = prompt,
            .Location = New Point(10, y),
            .Size = New Size(560, 40),
            .Font = New Font("맑은 고딕", 9, FontStyle.Bold)
        }
        Me.Controls.Add(lblPrompt)
        y += 50

        ' CheckedListBox
        checkedListBox = New CheckedListBox With {
            .Location = New Point(10, y),
            .Size = New Size(560, 320),
            .CheckOnClick = True
        }

        For Each link In allLinks
            checkedListBox.Items.Add(link.Name, False)
        Next

        Me.Controls.Add(checkedListBox)
        y += 330

        ' 전체 선택/해제 버튼
        btnSelectAll = New Button With {
            .Text = "전체 선택",
            .Location = New Point(10, y),
            .Size = New Size(100, 30)
        }
        AddHandler btnSelectAll.Click, AddressOf BtnSelectAll_Click
        Me.Controls.Add(btnSelectAll)

        btnDeselectAll = New Button With {
            .Text = "전체 해제",
            .Location = New Point(120, y),
            .Size = New Size(100, 30)
        }
        AddHandler btnDeselectAll.Click, AddressOf BtnDeselectAll_Click
        Me.Controls.Add(btnDeselectAll)

        ' OK/Cancel 버튼
        btnOK = New Button With {
            .Text = "확인",
            .Location = New Point(370, y),
            .Size = New Size(100, 30),
            .DialogResult = DialogResult.OK
        }
        AddHandler btnOK.Click, AddressOf BtnOK_Click
        Me.Controls.Add(btnOK)

        btnCancel = New Button With {
            .Text = "취소",
            .Location = New Point(470, y),
            .Size = New Size(100, 30),
            .DialogResult = DialogResult.Cancel
        }
        Me.Controls.Add(btnCancel)

        Me.AcceptButton = btnOK
        Me.CancelButton = btnCancel
    End Sub

    Private Sub BtnSelectAll_Click(sender As Object, e As EventArgs)
        For i = 0 To checkedListBox.Items.Count - 1
            checkedListBox.SetItemChecked(i, True)
        Next
    End Sub

    Private Sub BtnDeselectAll_Click(sender As Object, e As EventArgs)
        For i = 0 To checkedListBox.Items.Count - 1
            checkedListBox.SetItemChecked(i, False)
        Next
    End Sub

    Private Sub BtnOK_Click(sender As Object, e As EventArgs)
        SelectedLinks.Clear()

        For i = 0 To checkedListBox.CheckedIndices.Count - 1
            Dim index = checkedListBox.CheckedIndices(i)
            SelectedLinks.Add(allLinks(index))
        Next

        If SelectedLinks.Count = 0 Then
            MessageBox.Show("최소 1개 이상의 링크를 선택해야 합니다.")
            Me.DialogResult = DialogResult.None
        End If
    End Sub
End Class
