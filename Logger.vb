Imports System.IO

Public Class Logger
    Private Shared logFilePath As String = ""
    Private Shared lockObj As New Object()

    Public Shared Sub Initialize()
        Try
            Dim logFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ColumnNumberMapper_Logs")

            If Not Directory.Exists(logFolder) Then
                Directory.CreateDirectory(logFolder)
            End If

            logFilePath = Path.Combine(logFolder, "Log_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt")

            Log("로거 초기화 완료")
            Log($"로그 파일: {logFilePath}")

        Catch ex As Exception
            ' 로그 초기화 실패 시 콘솔에만 출력
            Console.WriteLine("Logger 초기화 실패: " & ex.Message)
        End Try
    End Sub

    Public Shared Sub Log(message As String)
        Try
            SyncLock lockObj
                Dim logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}"

                If Not String.IsNullOrEmpty(logFilePath) Then
                    File.AppendAllText(logFilePath, logMessage & Environment.NewLine)
                End If

                ' 콘솔에도 출력
                Console.WriteLine(logMessage)
            End SyncLock
        Catch ex As Exception
            Console.WriteLine("Log 쓰기 실패: " & ex.Message)
        End Try
    End Sub

    Public Shared Sub LogError(message As String)
        Try
            SyncLock lockObj
                Dim logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] *** ERROR *** {message}"

                If Not String.IsNullOrEmpty(logFilePath) Then
                    File.AppendAllText(logFilePath, logMessage & Environment.NewLine)
                End If

                ' 콘솔에도 출력
                Console.WriteLine(logMessage)
            End SyncLock
        Catch ex As Exception
            Console.WriteLine("Log 쓰기 실패: " & ex.Message)
        End Try
    End Sub

    Public Shared Function GetLogFilePath() As String
        Return logFilePath
    End Function
End Class