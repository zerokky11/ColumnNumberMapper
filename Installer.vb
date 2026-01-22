Imports System.IO
Imports System.Reflection

Public Class Installer

    Public Shared Sub Main(args() As String)
        ' 명령줄 인수로 /install이 전달되면 설치 수행
        If args.Length > 0 AndAlso args(0) = "/install" Then
            InstallAddin()
        End If
    End Sub

    Public Shared Sub InstallAddin()
        Try
            Console.WriteLine("=== Column Number Mapper 설치 시작 ===")

            ' Revit 2019 Addins 경로
            Dim addinsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                         "Autodesk", "Revit", "Addins", "2019")

            If Not Directory.Exists(addinsPath) Then
                Directory.CreateDirectory(addinsPath)
                Console.WriteLine($"Addins 폴더 생성: {addinsPath}")
            End If

            ' DLL 파일 경로
            Dim assemblyPath = Assembly.GetExecutingAssembly().Location
            Dim dllFileName = Path.GetFileName(assemblyPath)
            Dim targetDllPath = Path.Combine(addinsPath, dllFileName)

            ' DLL 복사 (빌드 후 자동 복사)
            If assemblyPath <> targetDllPath Then
                File.Copy(assemblyPath, targetDllPath, True)
                Console.WriteLine($"DLL 복사 완료: {targetDllPath}")
            End If

            ' .addin 파일 생성
            Dim addinFilePath = Path.Combine(addinsPath, "ColumnNumberMapper.addin")
            Dim addinContent = CreateAddinContent(targetDllPath)

            File.WriteAllText(addinFilePath, addinContent)
            Console.WriteLine($".addin 파일 생성 완료: {addinFilePath}")

            Console.WriteLine("=== 설치 완료 ===")
            Console.WriteLine("Revit 2019를 재시작하면 Add-Ins 탭에서 'Column Number Mapper'를 사용할 수 있습니다.")

        Catch ex As Exception
            Console.WriteLine($"설치 오류: {ex.Message}")
            Console.WriteLine(ex.StackTrace)
        End Try
    End Sub

    Private Shared Function CreateAddinContent(dllPath As String) As String
        Dim addinXml = $"<?xml version=""1.0"" encoding=""utf-8""?>
<RevitAddIns>
  <AddIn Type=""Command"">
    <Name>Column Number Mapper</Name>
    <Assembly>{dllPath}</Assembly>
    <AddInId>B5F5E3A7-9C2D-4F8B-A1E3-D6C4B7E8F9A0</AddInId>
    <FullClassName>ColumnNumberMapper</FullClassName>
    <Text>Column Number Mapper</Text>
    <Description>기둥 넘버를 구조 기둥의 S5_EQCODE 파라미터로 매핑하는 도구</Description>
    <VisibilityMode>AlwaysVisible</VisibilityMode>
    <VendorId>SMOO</VendorId>
    <VendorDescription>SAMOO Architects &amp; Engineers</VendorDescription>
  </AddIn>
</RevitAddIns>"

        Return addinXml
    End Function
End Class