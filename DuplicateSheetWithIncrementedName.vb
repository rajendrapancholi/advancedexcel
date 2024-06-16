Sub DuplicateSheetWithIncrementedName()
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim sheetCount As Integer
    Dim newName As String
    
    ' Get the current active sheet
    Set ws = ActiveSheet
    
    ' Count existing sheets
    sheetCount = Sheets.Count
    
    ' Create a new sheet as a copy of the active sheet
    ws.Copy After:=Sheets(sheetCount)
    
    ' Set the newly created sheet as the active sheet
    Set newSheet = ActiveSheet
    
    ' Modify the new sheet's name with the prefix and incremented number
    newName = "RAJE24" & sheetCount + 1
    newSheet.Name = newName
    
    ' Clean up
    Set ws = Nothing
    Set newSheet = Nothing
End Sub

