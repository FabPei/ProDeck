Attribute VB_Name = "StretchObjects"
Sub StretchToRight()
    ' Ensure at least two shapes are selected
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes.", vbInformation, "Selection Error"
        Exit Sub
    End If

    Dim lastShape As Shape
    ' The last selected shape is the target
    Set lastShape = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count)

    Dim targetRight As Single
    ' Calculate the right edge of the last selected shape
    targetRight = lastShape.Left + lastShape.Width

    Dim shp As Shape
    ' Loop through all selected shapes
    For Each shp In ActiveWindow.Selection.ShapeRange
        ' Exclude the last selected shape from being modified
        If shp.Name <> lastShape.Name Then
            ' Adjust the width of the shape to stretch to the target right edge
            shp.Width = targetRight - shp.Left
        End If
    Next shp
End Sub
Sub StretchToLeft()
    ' Ensure at least two shapes are selected
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes.", vbInformation, "Selection Error"
        Exit Sub
    End If

    Dim lastShape As Shape
    ' The last selected shape is the target
    Set lastShape = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count)

    Dim targetLeft As Single
    ' Get the left edge of the last selected shape
    targetLeft = lastShape.Left

    Dim shp As Shape
    ' Loop through all selected shapes
    For Each shp In ActiveWindow.Selection.ShapeRange
        ' Exclude the last selected shape from being modified
        If shp.Name <> lastShape.Name Then
            ' Adjust the width and left position of the shape
            shp.Width = (shp.Left + shp.Width) - targetLeft
            shp.Left = targetLeft
        End If
    Next shp
End Sub
Sub StretchToTop()
    ' Ensure at least two shapes are selected
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes.", vbInformation, "Selection Error"
        Exit Sub
    End If

    Dim lastShape As Shape
    ' The last selected shape is the target
    Set lastShape = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count)

    Dim targetTop As Single
    ' Get the top edge of the last selected shape
    targetTop = lastShape.Top

    Dim shp As Shape
    ' Loop through all selected shapes
    For Each shp In ActiveWindow.Selection.ShapeRange
        ' Exclude the last selected shape from being modified
        If shp.Name <> lastShape.Name Then
            ' Adjust the height and top position of the shape
            shp.Height = (shp.Top + shp.Height) - targetTop
            shp.Top = targetTop
        End If
    Next shp
End Sub
Sub StretchToBottom()
    ' Ensure at least two shapes are selected
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        MsgBox "Please select at least two shapes.", vbInformation, "Selection Error"
        Exit Sub
    End If

    Dim lastShape As Shape
    ' The last selected shape is the target
    Set lastShape = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count)

    Dim targetBottom As Single
    ' Calculate the bottom edge of the last selected shape
    targetBottom = lastShape.Top + lastShape.Height

    Dim shp As Shape
    ' Loop through all selected shapes
    For Each shp In ActiveWindow.Selection.ShapeRange
        ' Exclude the last selected shape from being modified
        If shp.Name <> lastShape.Name Then
            ' Adjust the height of the shape to stretch to the target bottom edge
            shp.Height = targetBottom - shp.Top
        End If
    Next shp
End Sub

