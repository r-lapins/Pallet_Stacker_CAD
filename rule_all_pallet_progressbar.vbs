Class ThisRule
    Private oProgressBar As Inventor.ProgressBar
    Private Property UserClickedOnCancel() As Boolean = False

    Public Sub Main()
        Dim folderPath As String = "C:\Users\ROBERT\Desktop\VELUX\Komponenty\Velux Pallets\Before the painting"
        Dim folder As Object
        Dim fs As Object = CreateObject("Scripting.FileSystemObject")

        Try
            For Each folder In fs.GetFolder(folderPath).SubFolders
                folder.Delete()
            Next folder
        Catch ex As Exception
            LogMessage("Błąd podczas usuwania folderów: " & ex.Message)
        End Try

        fs = Nothing
        folder = Nothing

        iLogicVb.UpdateWhenDone = True
        LogMessage("Start")

        Dim checkFileName As String = String.Empty
        Dim oApp As Inventor.Application = ThisApplication
        Dim oAsmDoc As AssemblyDocument = oApp.ActiveDocument
        Dim oCompDef As AssemblyComponentDefinition = oAsmDoc.ComponentDefinition
        Dim oTrans As Matrix
        Dim oFileSystem As Object = CreateObject("Scripting.FileSystemObject")
        Dim oFolder As Object = oFileSystem.GetFolder(folderPath)
        Dim oFile As Object
        Dim oOcc As ComponentOccurrence
        Dim xOffset As Double = 200
        Dim zOffset As Double = 400
        Dim currentX As Double = 0
        Dim currentZ As Double = -zOffset
        Dim i As Integer = 0
        Dim totalFiles As Integer = oFolder.Files.Count
        Dim fileCount As Integer = 0

        oProgressBar = ThisApplication.CreateProgressBar(False, totalFiles, "Porządek Palet w trakcie budowy...", True)
        AddHandler oProgressBar.OnCancel, AddressOf OnCancel
        UpDateProgress("Start")
        System.Threading.Tasks.Task.Delay(1000).Wait()
        If UserWantsToCancel() Then Return

        For Each oFile In oFolder.Files
            fileCount += 1
            Dim progressPercentage As Integer = CInt((fileCount / totalFiles) * 100)
            UpDateProgress("Przetwarzanie pliku nr: " & fileCount & " --> %" & progressPercentage)

            If UserWantsToCancel() Then Return

            If LCase(oFile.Name) Like "*.ipt" Or LCase(oFile.Name) Like "*.iam" Then
                If oFile.Name.Length >= 22 Then
                    If LCase(oFile.Name).Substring(11, 11) <> checkFileName Then
                        currentZ = currentZ + zOffset
                        currentX = 0
                        checkFileName = LCase(oFile.Name).Substring(11, 11)
                    End If
                Else
                    LogMessage("Nazwa pliku za krótka: " & oFile.Name)
                End If

                Try
                    If Not oFileSystem.FileExists(oFile.Path) Then
                        LogMessage("Plik nie istnieje: " & oFile.Path)
                    Else
                        oTrans = oApp.TransientGeometry.CreateMatrix
                        oTrans.SetTranslation(oApp.TransientGeometry.CreateVector(currentX, 0, currentZ))
                        oOcc = oCompDef.Occurrences.Add(oFile.Path, oTrans)
                        currentX = currentX + xOffset
                    End If
                Catch ex As Exception
                    LogMessage("Błąd podczas dodawania wystąpienia dla pliku: " & oFile.Name & " - " & ex.Message)
                    Exit Sub
                End Try
            End If
        Next oFile

        UpDateProgress("Saving assembly")
        If UserWantsToCancel() Then Return

        Dim sFilePath As String = ThisDoc.Path & "\"
        oAsmDoc.SaveAs(sFilePath & "All_Pallet_BTP.iam", False)

        LogMessage("Koniec")
        UpDateProgress("Complete")
        oProgressBar.Close()
    End Sub

    Private Sub UpDateProgress(message As String)
        oProgressBar.Message = message
        oProgressBar.UpdateProgress()
    End Sub

    Private Function UserWantsToCancel() As Boolean
        If (UserClickedOnCancel) Then
            MsgBox("Operation cancelled by user.")
            LogMessage("Operation cancelled by user.")
            oProgressBar.Close()
            Return True
        End If
        Return False
    End Function

    Private Sub OnCancel()
        UserClickedOnCancel = True
    End Sub

    'Logowanie błędów'
    Private Sub LogMessage(message As String)
        Dim logFilePath As String = "C:\Users\ROBERT\Desktop\VELUX\Komponenty\Velux Pallets\TransformAndReplaceLog.txt"
        Try
            Using writer As New System.IO.StreamWriter(logFilePath, True)
                writer.WriteLine(DateTime.Now.ToString("HH:mm:ss") & " - " & message)
            End Using
        Catch ex As Exception
            MessageBox.Show("Błąd zapisywania logu: " & ex.Message)
        End Try
    End Sub

End Class