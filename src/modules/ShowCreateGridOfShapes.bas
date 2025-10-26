Attribute VB_Name = "ShowCreateGridOfShapes"
Sub ShowSplitShapeWithForm()
    ' 1. Ensure a single shape is selected
    If ActiveWindow.Selection.ShapeRange.Count <> 1 Then
        MsgBox "Please select exactly one shape to split.", vbInformation
        Exit Sub
    End If

    ' 2. Create an instance of the form
    Dim form As Object
    Set form = New ufCreateGridOfShapes

    ' 3. Show the form and wait for the user
    ' The form will now handle all logic internally
    form.Show
    
    ' 4. Clean up after the form is closed
    Unload form
    Set form = Nothing
End Sub
