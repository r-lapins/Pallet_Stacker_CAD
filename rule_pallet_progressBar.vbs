Class ThisRule
    Private oProgressBar As Inventor.ProgressBar
    Private Property UserClickedOnCancel() As Boolean = False

    Public Sub Main()
        iLogicVb.UpdateWhenDone = True
        GoExcel.Open("C:\Users\ROBERT\Desktop\VELUX\Komponenty\velux_rule.xlsx", "Arkusz2")
        
        Dim Pocz As Integer = 5
        Dim Koniec As Integer = 186
        Dim sSpacerFolderPath As String = "C:\Users\ROBERT\Desktop\VELUX\Komponenty\Przekladki i palety\"
        Dim sPalletFolderPath As String = "C:\Users\ROBERT\Desktop\VELUX\Komponenty\Velux Pallets\Before the painting\"
        Dim sEPFolderPath As String = "C:\Users\ROBERT\Desktop\VELUX\Komponenty\Przekladki i palety\Paleta europejska\"
        Dim sCompFolderPath As String = "C:\Users\ROBERT\Desktop\VELUX\Komponenty\Velux Models_BTP\Models\"
        Dim componentName As String = "GDL P1 Bottom Sash MK:1"
        Dim FirstSpacerName As String = "S1_:1"
        Dim spacerName As String = "S3_:1"
        Dim EPName As String = "EP_800x1200x150:1"
        Dim NewEPName As String
        Dim PalletName As String
        Dim PalletOrient As String
        Dim NewPalletName As String
        Dim Spacer As String
        Dim SpacerLength As Double
        Dim SpacerHeight As Double
        Dim SpacerWidth As Double
        Dim FirstCompX As Double
        Dim FirstSpacerZ As Double
        Dim nComp_X As Double
        Dim nComp_Y As Double
        Dim nComp_X_dist As Double
        Dim CompHeight As Double
        Dim CompLength As Double
        Dim nSpacer_Y As Double
        Dim nSpacer_Z_dist As Double
        Dim PalletLength As Integer
        Dim FirstSpacerLength As Integer
        Dim SkipLoop As Double
        Dim oPattern As Object = Nothing
        Dim doc As Document = ThisDoc.Document
        Dim newComponentName As String
        Dim totalFiles As Integer = Koniec - Pocz - 1
        Dim fileCount As Integer = 0
        Dim CzopGora As Double
        Dim CzopDol As Double
        Dim CzopSkrYdist As Double
        Dim ObnizKomp As Integer

        oProgressBar = ThisApplication.CreateProgressBar(False, totalFiles, "Pędzenie palet pełnych produktów...", True)
        AddHandler oProgressBar.OnCancel, AddressOf OnCancel
        UpDateProgress("Start")
        System.Threading.Tasks.Task.Delay(1000).Wait()
        If UserWantsToCancel() Then Return

    For i = Pocz To Koniec
        
        fileCount += 1
        Dim progressPercentage As Integer = CInt((fileCount / totalFiles) * 100)
        UpDateProgress("Processing file: " & fileCount & " --> % " & progressPercentage)
        If UserWantsToCancel() Then Return

        LogMessage("Start: " & i)

        Try 
            SkipLoop = GoExcel.CellValue("A" & i)
        Catch ex As Exception
            GoTo ContinueLoop
        End Try
        If SkipLoop <> 1 And SkipLoop <> 2 Then GoTo ContinueLoop

        newComponentName = GoExcel.CellValue("B" & i) + " " + GoExcel.CellValue("C" & i) + " " + GoExcel.CellValue("D" & i)
        PalletName = "Pallet_BTP " + newComponentName
        FirstCompX = GoExcel.CellValue("T" & i)
        nComp_X = GoExcel.CellValue("O" & i)
        nComp_Y = GoExcel.CellValue("P" & i)
        nComp_X_dist = GoExcel.CellValue("E" & i)
        CompHeight = GoExcel.CellValue("F" & i)
        CompLength = GoExcel.CellValue("G" & i)
        Spacer = GoExcel.CellValue("V" & i)
        SpacerLength = GoExcel.CellValue("U" & i)
        SpacerHeight = GoExcel.CellValue("W" & i)
        SpacerWidth = GoExcel.CellValue("X" & i)
        nSpacer_Y = GoExcel.CellValue("P" & i)
        nSpacer_Z_dist = GoExcel.CellValue("Y" & i)
        PalletOrient = GoExcel.CellValue("M" & i)
        PalletLength = GoExcel.CellValue("I" & i)
        FirstSpacerZ = GoExcel.CellValue("Z" & i)
        NewEPName = "EP_800x" & PalletLength & "x150"
        CzopGora = GoExcel.CellValue("K" & i)
        CzopDol = GoExcel.CellValue("L" & i)
        CzopSkrYdist = CzopGora + CzopDol
        ObnizKomp = GoExcel.CellValue("AB" & i)

        If PalletOrient = "Lengthwise" Then
            FirstSpacerLength = 800
            Parameter("PalletAngle") = 90
        ElseIf PalletOrient = "Crosswise" Then
            FirstSpacerLength = PalletLength
            Parameter("PalletAngle") = 0
        End If
        
        Parameter("n_Comp_X_dist") = nComp_X_dist
        Parameter("n_Comp_Y_dist") = CompHeight + SpacerHeight - CzopSkrYdist
        Parameter("n_Comp_X") = nComp_X
        Parameter("n_Comp_Y") = nComp_Y
        Parameter("n_Spacer_Y") = nSpacer_Y
        Parameter("n_Spacer_Y_dist") = CompHeight + SpacerHeight - CzopSkrYdist
        Parameter("n_Spacer_Z_dist") = nSpacer_Z_dist
        Parameter("FirstSpacerGap") = FirstSpacerZ * 2

        Dim newXOffset As Double = FirstCompX / 10.0
        Dim newYOffset As Double = (150 + 10) / 10.0
        If ObnizKomp = 1 Then newYOffset = newYOffset - CzopDol / 10.0
        Dim newZOffset As Double = 0.0
        Dim newXOffset_S As Double = 0
        Dim newYOffset_S As Double = newYOffset + (CompHeight + SpacerHeight / 2 - CzopGora) / 10.0
        Dim newZOffset_S As Double = nSpacer_Z_dist / 2 / 10.0
        Dim newXOffset_FS As Double = 0
        Dim newYOffset_FS As Double = 155 / 10.0
        Dim newZOffset_FS As Double = FirstSpacerZ / 10.0

        ' Pobierz aktywny dokument, złożenia, wsio wystąpienia komponentów w złożeniach
        Dim oAsmDoc As AssemblyDocument = ThisApplication.ActiveDocument
        Dim oCompDef As AssemblyComponentDefinition = oAsmDoc.ComponentDefinition
        Dim oOccurrences As ComponentOccurrences = oCompDef.Occurrences
        Dim oOccurrence As ComponentOccurrence

        ' Europaleta
        TransformAndReplaceComponent(oOccurrences, EPName, NewEPName, sEPFolderPath, 0, 0, 0)

        ' Pierwsza przekładka
        Dim newFirstSpacerName As String = "S1_" & FirstSpacerLength & "_f"
        TransformAndReplaceComponent(oOccurrences, FirstSpacerName, newFirstSpacerName, sSpacerFolderPath, newXOffset_FS, newYOffset_FS, newZOffset_FS)

        ' Profil
        TransformAndReplaceComponent(oOccurrences, componentName, newComponentName, sCompFolderPath, newXOffset, newYOffset, newZOffset)
        
        ' Otwórz przekładkę, zmodyfikuj, zapisz jako nowy plik, podmień
        Dim newSpacerName As String = Spacer & "_" & SpacerLength
        ' Sprawdz czy jest taka przekładka - "nie ma" -> stwórz takową
        If Dir(sSpacerFolderPath & newSpacerName & ".ipt") = "" Then
            ' Podmiana przekladki
            Dim filePath As String = sSpacerFolderPath & Spacer & "_.ipt"
            Dim oPartDoc As Document = ThisApplication.Documents.Open(filePath)
            Dim oPartDef As PartComponentDefinition = oPartDoc.ComponentDefinition
            oPartDef.Parameters.Item("Length").Value = SpacerLength / 10
            oPartDoc.SaveAs(sSpacerFolderPath & newSpacerName & ".ipt", False)
            oPartDoc.Close(True)
        End If
        TransformAndReplaceComponent(oOccurrences, spacerName, newSpacerName, sSpacerFolderPath, newXOffset_S, newYOffset_S, newZOffset_S)

        UpdateAndSave(oPattern, doc, sPalletFolderPath, PalletName)
        
        ' Dla wersji Right/Left
        If SkipLoop = 2 Then
            newComponentName = newComponentName + "_Left"
            PalletName = "Pallet_BTP " + newComponentName
            TransformAndReplaceComponent(oOccurrences, componentName, newComponentName, sCompFolderPath, newXOffset, newYOffset, newZOffset)
            UpdateAndSave(oPattern, doc, sPalletFolderPath, PalletName)
        End If

        ContinueLoop:
        Next i
        LogMessage("END")

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

    Private Sub UpdateAndSave(ByVal oPattern As Object, ByVal doc As Document, ByVal sPalletFolderPath As String, ByVal PalletName As String)
        If Not oPattern Is Nothing Then
            oPattern.Update()
        End If
        doc.SaveAs(sPalletFolderPath & PalletName & ".iam", False)
    End Sub

    Private Sub TransformAndReplaceComponent(ByRef oOccurrences As ComponentOccurrences, ByRef componentName As String, ByVal newComponentName As String, ByVal folderPath As String, ByVal xOffset As Double, ByVal yOffset As Double, ByVal zOffset As Double)
        Dim oOccurrence As ComponentOccurrence = Nothing
        Dim found As Boolean = False

        ' Wyszukiwanie komponentu
        For Each oOccurrence In oOccurrences
            If oOccurrence.Name = componentName Then
                found = True
                Exit For
            End If
        Next

        If Not found Then
            LogMessage("Komponent " & componentName & " nie został znaleziony w bieżącej iteracji.")
            Exit Sub
        End If

        ' Transformacja komponentu
        Dim oTransform As Matrix = oOccurrence.Transformation
        oTransform.Cell(1, 4) = xOffset
        oTransform.Cell(2, 4) = yOffset
        oTransform.Cell(3, 4) = zOffset
        oOccurrence.Transformation = oTransform
        
        ' Podmiana komponentu
        Dim sFilePath As String = folderPath & newComponentName & ".ipt"
        If Not System.IO.File.Exists(sFilePath) Then
            LogMessage("Plik do podmiany nie istnieje: " & sFilePath)
            Exit Sub
        End If
        oOccurrence.Replace(sFilePath, True)
        componentName = newComponentName & ":1"
    End Sub

    'Logowanie błędów'
    Private Sub LogMessage(message As String)
        Dim logFilePath As String =  "C:\Users\ROBERT\Desktop\VELUX\Komponenty\Velux Pallets\TransformAndReplaceLog.txt"
        Try
            Using writer As New System.IO.StreamWriter(logFilePath, True)
                writer.WriteLine(DateTime.Now.ToString("HH:mm:ss") & " - " & message)
            End Using
        Catch ex As Exception
            MessageBox.Show("Błąd zapisywania logu: " & ex.Message)
        End Try
    End Sub
End Class