VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AddNewRecipe()
    Dim wsTemplate As Worksheet
    Dim wsIndex As Worksheet
    Dim newRecipeSheet As Worksheet
    Dim recipeName As String
    Dim newSheetName As String
    Dim nextRow As Long

    ' set references to the template and index sheets
    On Error Resume Next
    Set wsTemplate = ThisWorkbook.Sheets("ADD_RECIPES")
    Set wsIndex = ThisWorkbook.Sheets("INDEX")
    On Error GoTo 0

    ' debugging: Print the sheet names to Immediate Window
    If wsTemplate Is Nothing Then
        MsgBox "Template sheet not found", vbCritical
        Exit Sub
    End If
    If wsIndex Is Nothing Then
        MsgBox "Index sheet not found", vbCritical
        Exit Sub
    End If
    Debug.Print "Template Sheet: " & wsTemplate.Name
    Debug.Print "Index Sheet: " & wsIndex.Name

    ' get the recipe name from the template
    recipeName = wsTemplate.Range("C4").Value
    If recipeName = "" Then
        MsgBox "Recipe name cannot be empty", vbExclamation
        Exit Sub
    End If
    Debug.Print "Recipe Name: " & recipeName

    ' generate a new sheet name based on the recipe name
    newSheetName = "Recipe_" & recipeName
    Debug.Print "New Sheet Name: " & newSheetName

    ' copy the template sheet to create a new sheet
    On Error Resume Next
    wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set newRecipeSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    On Error GoTo 0
    If newRecipeSheet Is Nothing Then
        MsgBox "Error creating new sheet", vbCritical
        Exit Sub
    End If
    newRecipeSheet.Name = newSheetName
    Debug.Print "New Recipe Sheet: " & newRecipeSheet.Name

    ' update the index sheet with the new recipe name and hyperlink
    On Error GoTo ErrorHandler
    nextRow = wsIndex.Cells(wsIndex.Rows.Count, 1).End(xlUp).Row + 1
    Debug.Print "Next Row: " & nextRow
    Debug.Print "Adding Hyperlink to: " & newSheetName & "!C4"
    wsIndex.Hyperlinks.Add Anchor:=wsIndex.Cells(nextRow, 1), Address:="", SubAddress:= _
        "'" & newSheetName & "'!C4", TextToDisplay:=recipeName

    ' clear only specific cells in the template
    wsTemplate.Range("C4").Value = ""
    wsTemplate.Range("H4").Value = ""
    wsTemplate.Range("B8:I23").ClearContents
    wsTemplate.Range("B27:B41").ClearContents
    wsTemplate.Range("B45:B51").ClearContents
    MsgBox "Recipe added successfully!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub


