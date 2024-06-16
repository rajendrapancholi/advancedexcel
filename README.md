# To create a macro in Excel that duplicates the current sheet, adds a prefix "RAJE24" followed by an incremented number to the new sheet's name, you can follow these steps:

## 1. Open the Visual Basic for Applications (VBA) Editor:

- Press Alt + F11 in Excel to open the VBA Editor.

## 2. Insert a New Module:

- In the VBA Editor, go to Insert > Module to insert a new module where you'll write your macro.

## 3. Write the Macro:

- Copy and paste the following VBA code into the module window:

```vb
Copy code
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
    newName = "RAJE24 " & sheetCount + 1
    newSheet.Name = newName

    ' Clean up
    Set ws = Nothing
    Set newSheet = Nothing

End Sub
```

## 4. Save the Macro:

- Close the VBA Editor.
- Your macro is now saved in the workbook. You can run it whenever needed.

## 5. Run the Macro:

- Go back to Excel.
- Press Alt + F8 to open the "Macro" dialog box.
- Select DuplicateSheetWithIncrementedName and click Run.

## Explanation of the Macro:

- Sub DuplicateSheetWithIncrementedName(): This is the main subroutine that executes when you run the macro.

- Dim statements: These declare variables used in the macro.

- ws: This variable represents the current active worksheet.

- newSheet: This variable will represent the newly created worksheet.

- sheetCount: This variable holds the count of existing sheets.

- ws.Copy After:=Sheets(sheetCount): This line duplicates the active sheet after the last existing sheet.
- newSheet: This sets the newly created sheet as the active sheet.

- newName = "RAJE24" & sheetCount + 1: This constructs the new name with the prefix "RAJE24" followed by an incremented number (sheetCount + 1).

- newSheet.Name = newName: This renames the newly created sheet with the constructed name.

- Clean up: Finally, it clears the variables to release memory.

#### This macro will duplicate the current active sheet and rename the duplicate with the prefix "RAJE24" followed by an incremented number (starting from 1 if there are no existing sheets). Adjust the prefix or the renaming logic (newName assignment) as per your specific requirements.
