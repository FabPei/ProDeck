Attribute VB_Name = "ScaleSelectionRelative"
Option Explicit

' --- Solid Integer Constants (to avoid 'library not found' errors) ---
Const MsoTriState_True As Long = -1         ' msoTrue

' RelativeToOriginalSize: msoFalse/0 means scale from the *current* size.
Const MsoScaleRelative_Current As Long = 0
' -----------------------------------------------------------------------------

'================================================================================
' Macro:        ScaleSelectionRelative
' Purpose:      Scales selected shapes and their internal font sizes.
'               It manually simulates "Scale from Middle" to bypass VBA argument issues.
'================================================================================
Sub ScaleSelectionRelative()

    ' --- Variable Declarations ---
    Dim sel As Selection
    Dim shpRange As ShapeRange
    Dim sInput As String
    Dim scaleFactor As Single
    
    ' We need a single object to work on, even if it's a temporary group
    Dim shp As Shape
    Dim groupedShape As Shape
    Dim individualShape As Shape ' For iterating over the original shapes
    
    ' Variables for our manual "scale from middle" math
    Dim OldTop As Single, OldLeft As Single
    Dim OldHeight As Single, OldWidth As Single
    Dim NewTop As Single, NewLeft As Single
    Dim NewHeight As Single, NewWidth As Single
    Dim CenterX As Single, CenterY As Single
    
    On Error GoTo ErrorHandler

    ' 1. Get the active selection
    If Not ActiveWindow Is Nothing Then
        Set sel = ActiveWindow.Selection
    Else
        Exit Sub
    End If

    ' 2. Validation checks
    If sel.Type <> ppSelectionShapes Or sel.ShapeRange.Count = 0 Then
        MsgBox "Please select one or more objects to scale.", vbInformation, "Nothing Selected"
        Exit Sub
    End If

    Set shpRange = sel.ShapeRange

    ' 3. Get user input
    sInput = InputBox("Enter the scaling factor in percent:" & vbCrLf & _
                      "(e.g., 120 for 120% or 80 for 80%)", _
                      "Scale Objects", "100")
                      
    If sInput = "" Then Exit Sub ' User pressed Cancel
    
    If Not IsNumeric(sInput) Then
        MsgBox "That doesn't look like a number. Please enter a valid percentage (e.g., 120).", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    scaleFactor = CSng(sInput) / 100#  ' Convert "120" to 1.2
    
    If scaleFactor <= 0 Then
        MsgBox "The scaling factor must be a positive number (greater than 0).", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    ' 4. Set up the object reference for scaling
    
    If shpRange.Count = 1 Then
        ' --- Case 1: Just one object is selected ---
        Set shp = shpRange(1)
    Else
        ' --- Case 2: Multiple objects (we'll group them temporarily) ---
        Set groupedShape = shpRange.Group
        Set shp = groupedShape ' Now 'shp' refers to the whole group
    End If
    
    ' --- Manual "Scale from Middle" Simulation (Geometry Scaling) ---

    ' 1. Store the shape's current position and size
    OldTop = shp.Top
    OldLeft = shp.Left
    OldHeight = shp.Height
    OldWidth = shp.Width

    ' 2. Calculate the exact center point
    CenterX = OldLeft + (OldWidth / 2)
    CenterY = OldTop + (OldHeight / 2)

    ' 3. Lock the aspect ratio (so it doesn't get stretched)
    shp.LockAspectRatio = MsoTriState_True

    ' 4. Scale the shape (this defaults to scaling from the top-left corner)
    '    .ScaleHeight needs two arguments. 0 (MsoScaleRelative_Current) works for all types.
    shp.ScaleHeight Factor:=scaleFactor, _
                    RelativeToOriginalSize:=MsoScaleRelative_Current

    ' 5. Get the new height and width (after scaling)
    NewHeight = shp.Height
    NewWidth = shp.Width

    ' 6. Calculate the new Top/Left position to keep the center point the same
    NewTop = CenterY - (NewHeight / 2)
    NewLeft = CenterX - (NewWidth / 2)

    ' 7. Move the shape to its new position, completing the "middle" effect
    shp.Top = NewTop
    shp.Left = NewLeft
    
    ' --- End of Geometry Scaling ---

    ' 5. Apply font scaling to all individual shapes in the original range
    For Each individualShape In shpRange
        If individualShape.HasTextFrame Then
            Dim txtRange As TextRange
            Set txtRange = individualShape.TextFrame.TextRange
            
            ' Check if there is actual text and a font size to scale
            If txtRange.Length > 0 Then
                ' Scale the font size by the same factor
                txtRange.Font.Size = txtRange.Font.Size * scaleFactor
            End If
        End If
    Next individualShape

    ' 6. Clean up
    If Not groupedShape Is Nothing Then
        groupedShape.Ungroup
    End If

    Exit Sub

' --- Simple Error Handling ---
ErrorHandler:
    MsgBox "An unexpected error occurred: " & vbCrLf & Err.Description, vbCritical, "Error During Scaling"
End Sub

