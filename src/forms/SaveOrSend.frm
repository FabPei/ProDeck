VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SaveOrSend 
   Caption         =   "Save or Send Presentation or selected Slides"
   ClientHeight    =   7488
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   9808.001
   OleObjectBlob   =   "SaveOrSend.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SaveOrSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonOk_Click()
    ' --- Validation Step ---
    If CheckBoxSlidesAll.Value = False And CheckBoxSlidesSelected.Value = False And CheckBoxSlidesVisible.Value = False Then
        MsgBox "Please select a slide range option (All, Selected, or Visible).", vbExclamation, "Selection Missing"
        Exit Sub
    End If
    
    If CheckBoxpptx.Value = False And CheckBoxpdf.Value = False And CheckBoxpdfpptx.Value = False And CheckBoxpdfprotected.Value = False Then
        MsgBox "Please select a file format option.", vbExclamation, "Selection Missing"
        Exit Sub
    End If

    If OptionButtonFile.Value = False And OptionButtonClipboard.Value = False And OptionButtonEmail.Value = False Then
        MsgBox "Please select an output action (File, Clipboard, or Email).", vbExclamation, "Selection Missing"
        Exit Sub
    End If

    ' --- Main processing starts here ---
    Dim pptPres As Presentation
    Set pptPres = ActivePresentation
    
    Dim newPres As Presentation
    Set newPres = Presentations.Add(msoFalse) ' Create a new, invisible presentation

    ' --- Step 1: Copy the correct slides to the new presentation ---
    If CheckBoxSlidesAll.Value = True Then
        pptPres.Slides.Range.Copy
        newPres.Slides.Paste
    ElseIf CheckBoxSlidesSelected.Value = True Then
        If ActiveWindow.Selection.Type <> ppSelectionSlides Then
            MsgBox "Please select one or more slides in the slide navigation pane.", vbExclamation, "No Slides Selected"
            newPres.Close
            Exit Sub
        End If
        ActiveWindow.Selection.SlideRange.Copy
        newPres.Slides.Paste
    ElseIf CheckBoxSlidesVisible.Value = True Then
        Dim sld As Slide
        For Each sld In pptPres.Slides
            If sld.SlideShowTransition.Hidden = msoFalse Then
                sld.Copy
                newPres.Slides.Paste (newPres.Slides.Count + 1)
            End If
        Next sld
    End If
    
    ' Validate that slides were actually copied
    If newPres.Slides.Count = 0 Then
         MsgBox "There were no slides to process based on your selection.", vbInformation, "No Slides Found"
         newPres.Close
         Exit Sub
    ElseIf newPres.Slides.Count > 1 Then
        newPres.Slides(1).Delete
    End If

    ' --- Step 2: Perform the selected output action ---

    ' Generate the base filename from the textboxes
    Dim baseFileName As String
    baseFileName = Format(CDate(Me.TextBoxDate.Value), "yyyymmdd") & "_" & Me.TextBoxTopic.Value
    
    ' -- Action: Save to File --
    If OptionButtonFile.Value = True Then
        Dim dialog As FileDialog
        Set dialog = Application.FileDialog(msoFileDialogSaveAs)
        
        With dialog
            .Title = "Save As"
            
            ' Set the initial filename in the dialog
            If CheckBoxpdfpptx.Value Then
                .initialFileName = baseFileName & ".pptx" ' Dialog saves the PPTX part
            Else
                .initialFileName = baseFileName & IIf(CheckBoxpptx.Value, ".pptx", ".pdf")
            End If
            
            ' Set the correct filter
            Dim i As Integer
            Dim filterExt As String
            filterExt = IIf(CheckBoxpptx.Value Or CheckBoxpdfpptx.Value, "*.pptx", "*.pdf")
            For i = 1 To .Filters.Count
                If LCase(.Filters(i).Extensions) = filterExt Then
                    .FilterIndex = i
                    Exit For
                End If
            Next i

            ' Show the dialog and process based on user action
            If .Show = -1 Then
                Dim savePath As String
                savePath = .SelectedItems(1)
                
                If CheckBoxpptx.Value Then
                    newPres.SaveAs savePath
                    MsgBox "Presentation successfully saved as PPTX.", vbInformation, "Success"
                
                ElseIf CheckBoxpdf.Value Then
                    newPres.ExportAsFixedFormat Path:=savePath, FixedFormatType:=ppFixedFormatTypePDF
                    MsgBox "Presentation successfully saved as PDF.", vbInformation, "Success"
                    
                ElseIf CheckBoxpdfprotected.Value Then
                    Dim pdfPassword As String
                    pdfPassword = InputBox("Please enter a password for the PDF:", "PDF Protection")
                    If pdfPassword <> "" Then
                        newPres.Password = pdfPassword
                        newPres.ExportAsFixedFormat Path:=savePath, FixedFormatType:=ppFixedFormatTypePDF
                        newPres.Password = "" ' Clear password from temp object
                        MsgBox "Protected PDF successfully saved.", vbInformation, "Success"
                    Else
                        MsgBox "Save cancelled. No password was provided.", vbExclamation, "Cancelled"
                    End If
                    
                ElseIf CheckBoxpdfpptx.Value Then
                    ' Save the PPTX from the dialog path
                    newPres.SaveAs savePath
                    ' Automatically save the PDF in the same location with the same base name
                    Dim pdfPath As String
                    pdfPath = Left(savePath, InStrRev(savePath, ".") - 1) & ".pdf"
                    newPres.ExportAsFixedFormat Path:=pdfPath, FixedFormatType:=ppFixedFormatTypePDF
                    MsgBox "PPTX and PDF files successfully saved in the same location.", vbInformation, "Success"
                End If
            End If
        End With

    ' -- Action: Copy File to Clipboard --
    ElseIf OptionButtonClipboard.Value = True Then
        ' This example will create both files in TEMP and copy both.
        
        If CheckBoxpdfpptx.Value Then
            ' Create both files
            Dim tempPptxPath As String, tempPdfPath As String
            tempPptxPath = Environ("TEMP") & "\" & baseFileName & ".pptx"
            tempPdfPath = Environ("TEMP") & "\" & baseFileName & ".pdf"
            newPres.SaveAs tempPptxPath
            newPres.ExportAsFixedFormat Path:=tempPdfPath, FixedFormatType:=ppFixedFormatTypePDF
            ' Note: Copying multiple files to clipboard is complex; this copies just the PPTX path.
            Dim dataObj As Object
            Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
            dataObj.SetText tempPptxPath
            dataObj.PutInClipboard
            MsgBox "PPTX and PDF created in TEMP folder. The PPTX file has been copied to the clipboard.", vbInformation, "Files Created"

        Else
            ' Handle single file copy
            Dim tempFilePath As String
            tempFilePath = Environ("TEMP") & "\" & baseFileName & IIf(CheckBoxpptx.Value, ".pptx", ".pdf")
            
            If CheckBoxpdfprotected.Value Then
                Dim clipPassword As String
                clipPassword = InputBox("Please enter a password for the PDF:", "PDF Protection")
                If clipPassword <> "" Then
                    newPres.Password = clipPassword
                    newPres.ExportAsFixedFormat Path:=tempFilePath, FixedFormatType:=ppFixedFormatTypePDF
                    newPres.Password = ""
                    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
                    dataObj.SetText tempFilePath
                    dataObj.PutInClipboard
                    MsgBox "The protected PDF has been created and copied to the clipboard.", vbInformation, "File Copied"
                Else
                    MsgBox "Copy cancelled. No password provided.", vbExclamation
                End If
            Else
                If CheckBoxpptx.Value Then
                    newPres.SaveAs tempFilePath
                Else
                    newPres.ExportAsFixedFormat Path:=tempFilePath, FixedFormatType:=ppFixedFormatTypePDF
                End If
                Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
                dataObj.SetText tempFilePath
                dataObj.PutInClipboard
                MsgBox "The file has been created and copied to the clipboard.", vbInformation, "File Copied"
            End If
        End If

    ' -- Action: Send as Email --
    ElseIf OptionButtonEmail.Value = True Then
        Dim OutApp As Object, OutMail As Object
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)
        
        With OutMail
            .Subject = Me.TextBoxTopic.Value
            .Body = "Please find the attached presentation(s)."
            
            If CheckBoxpdfpptx.Value Then
                ' Create and attach both files
                Dim emailPptxPath As String, emailPdfPath As String
                emailPptxPath = Environ("TEMP") & "\" & baseFileName & ".pptx"
                emailPdfPath = Environ("TEMP") & "\" & baseFileName & ".pdf"
                newPres.SaveAs emailPptxPath
                newPres.ExportAsFixedFormat Path:=emailPdfPath, FixedFormatType:=ppFixedFormatTypePDF
                .Attachments.Add emailPptxPath
                .Attachments.Add emailPdfPath
            Else
                ' Create and attach a single file
                Dim emailFilePath As String
                emailFilePath = Environ("TEMP") & "\" & baseFileName & IIf(CheckBoxpptx.Value, ".pptx", ".pdf")
                
                If CheckBoxpptx.Value Then
                    newPres.SaveAs emailFilePath
                ElseIf CheckBoxpdf.Value Then
                    newPres.ExportAsFixedFormat Path:=emailFilePath, FixedFormatType:=ppFixedFormatTypePDF
                ElseIf CheckBoxpdfprotected.Value Then
                    Dim emailPdfPassword As String
                    emailPdfPassword = InputBox("Please enter a password for the PDF attachment:", "PDF Protection")
                    If emailPdfPassword <> "" Then
                        newPres.Password = emailPdfPassword
                        newPres.ExportAsFixedFormat Path:=emailFilePath, FixedFormatType:=ppFixedFormatTypePDF
                        newPres.Password = ""
                    Else
                        MsgBox "Email cancelled. No password was provided.", vbExclamation, "Cancelled"
                        newPres.Close
                        Exit Sub
                    End If
                End If
                .Attachments.Add emailFilePath
            End If
            .Display
        End With
        
        Set OutMail = Nothing
        Set OutApp = Nothing
    End If

    ' --- Final Cleanup ---
    newPres.Close
    Unload Me ' Close the form after the operation is complete
End Sub


Private Sub UserForm_Initialize()
    ' Set the date and topic textboxes
    Me.TextBoxDate.Value = Format(Date, "dd.mm.yyyy")
    If ActivePresentation.Name <> "" Then
        Me.TextBoxTopic.Value = Left(ActivePresentation.Name, (InStrRev(ActivePresentation.Name, ".", -1, vbTextCompare) - 1))
    End If
    
    ' --- Uncheck all options at startup ---
    ' Uncheck all file format checkboxes
    Me.CheckBoxpptx.Value = False
    Me.CheckBoxpdf.Value = False
    Me.CheckBoxpdfpptx.Value = False
    Me.CheckBoxpdfprotected.Value = False
    
    ' Uncheck all slide range checkboxes
    Me.CheckBoxSlidesAll.Value = False
    Me.CheckBoxSlidesSelected.Value = False
    Me.CheckBoxSlidesVisible.Value = False
    
    ' Unselect all output action OptionButtons
    Me.OptionButtonFile.Value = False
    Me.OptionButtonClipboard.Value = False
    Me.OptionButtonEmail.Value = False
    
    ' --- NEW: Dynamically update checkbox captions ---
    Dim totalSlides As Integer
    Dim selectedSlides As Integer
    
    ' Get the total number of slides in the presentation
    totalSlides = ActivePresentation.Slides.Count
    
    ' Get the number of currently selected slides
    If ActiveWindow.Selection.Type = ppSelectionSlides Then
        selectedSlides = ActiveWindow.Selection.SlideRange.Count
    Else
        selectedSlides = 0
    End If
    
    ' Set the captions
    Me.CheckBoxSlidesAll.Caption = "Complete Presentation (" & totalSlides & " slides)"
    Me.CheckBoxSlidesSelected.Caption = "Selected Slides (" & selectedSlides & " slides)"
    ' --- END OF NEW CODE ---
    
    ' Call the update routine to set the initial file name
    UpdateFileName
End Sub



' This event runs whenever the date textbox is changed
Private Sub TextBoxDate_Change()
    UpdateFileName
End Sub

' This event runs whenever the topic textbox is changed
Private Sub TextBoxTopic_Change()
    UpdateFileName
End Sub

Private Sub UpdateFileName()
    Dim formattedDate As String
    Dim topic As String
    Dim fileExtension As String
    
    On Error Resume Next
    formattedDate = Format(CDate(Me.TextBoxDate.Value), "yyyymmdd")
    If Err.Number <> 0 Then
        formattedDate = Me.TextBoxDate.Value
        Err.Clear
    End If
    On Error GoTo 0
    
    topic = Me.TextBoxTopic.Value
    
    ' Determine the correct file extension for display
    If Me.CheckBoxpptx.Value = True Then
        fileExtension = ".pptx"
    ElseIf Me.CheckBoxpdf.Value = True Then
        fileExtension = ".pdf"
    ElseIf Me.CheckBoxpdfprotected.Value = True Then
        fileExtension = ".pdf"
    ElseIf Me.CheckBoxpdfpptx.Value = True Then
        fileExtension = " (.pptx + .pdf)" ' Special case
    End If
    
    Me.TextBoxFileName.Value = formattedDate & "_" & topic & fileExtension
End Sub



' Closes the UserForm when the Cancel button is clicked
Private Sub ButtonCancel_Click()
    Unload Me
End Sub
' --- Event handlers to make format checkboxes mutually exclusive ---
' --- CORRECTED: Event handlers for FORMAT checkboxes ---

Private Sub CheckBoxpptx_Click()
    If Me.CheckBoxpptx.Value = True Then
        Me.CheckBoxpdf.Value = False
        Me.CheckBoxpdfpptx.Value = False
        Me.CheckBoxpdfprotected.Value = False
    End If
    UpdateFileName
End Sub

Private Sub CheckBoxpdf_Click()
    If Me.CheckBoxpdf.Value = True Then
        Me.CheckBoxpptx.Value = False
        Me.CheckBoxpdfpptx.Value = False
        Me.CheckBoxpdfprotected.Value = False
    End If
    UpdateFileName
End Sub

Private Sub CheckBoxpdfpptx_Click()
    If Me.CheckBoxpdfpptx.Value = True Then
        Me.CheckBoxpptx.Value = False
        Me.CheckBoxpdf.Value = False
        Me.CheckBoxpdfprotected.Value = False
    End If
    UpdateFileName
End Sub

Private Sub CheckBoxpdfprotected_Click()
    If Me.CheckBoxpdfprotected.Value = True Then
        Me.CheckBoxpptx.Value = False
        Me.CheckBoxpdf.Value = False
        Me.CheckBoxpdfpptx.Value = False
    End If
    UpdateFileName
End Sub

' --- CORRECTED: Event handlers for RANGE checkboxes ---

Private Sub CheckBoxSlidesAll_Click()
    If Me.CheckBoxSlidesAll.Value = True Then
        Me.CheckBoxSlidesSelected.Value = False
        Me.CheckBoxSlidesVisible.Value = False
    End If
End Sub

Private Sub CheckBoxSlidesSelected_Click()
    If Me.CheckBoxSlidesSelected.Value = True Then
        Me.CheckBoxSlidesAll.Value = False
        Me.CheckBoxSlidesVisible.Value = False
    End If
End Sub

Private Sub CheckBoxSlidesVisible_Click()
    If Me.CheckBoxSlidesVisible.Value = True Then
        Me.CheckBoxSlidesAll.Value = False
        Me.CheckBoxSlidesSelected.Value = False
    End If
End Sub
