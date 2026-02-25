' =============================================================================
' ApplyTemplateStyles - Copy Headers/Footers + Black Text + Table Header Format
' =============================================================================
' Copies header/footer content, syncs header/footer font type, applies body
' font from template, forces text color to black, and applies the same table
' header formatting to all tables in the active document.
' Creates a timestamped backup before making changes.
'
' IMPORTANT: Open BOTH documents in Word before running:
'   1. Your target document (the one you want to format)
'   2. Your template document (the one with the correct formatting)
'
' Usage:
'   1. Open both documents in Word
'   2. Click into your TARGET document (make it the active window)
'   3. Run the macro (Alt+F8 > ApplyTemplateStyles > Run)
'   4. Select your template from the list of open documents
'   5. Review the summary dialog
'
' To install:
'   - Open Word > Alt+F11 > Insert > Module > Paste this code
'   - Or save into Normal.dotm for access from any document
' =============================================================================

Option Explicit

' Set to True to disable "auto update styles on open" after applying.
Private Const DISABLE_AUTO_UPDATE As Boolean = True

Public Sub ApplyTemplateStyles()

    Dim doc As Document
    Dim tmplDoc As Document
    Dim backupPath As String
    Dim headersFootersCopied As Boolean
    Dim headerFooterFontsSynced As Boolean
    Dim fontSetFromTemplate As Boolean
    Dim textColorSetToBlack As Boolean
    Dim templateFontName As String
    Dim tableHeadersFormatted As Long
    Dim tableHeadersUpdated As Boolean
    Dim startTime As Single
    Dim currentStep As String

    ' --- Guard: need at least 2 documents open ---
    If Documents.Count < 2 Then
        MsgBox "Please open BOTH your target document and your template document in Word before running this macro." & vbCrLf & vbCrLf & _
               "Currently open: " & Documents.Count & " document(s)", _
               vbExclamation, "Apply Template Styles"
        Exit Sub
    End If

    Set doc = ActiveDocument
    startTime = Timer

    ' --- Guard: document must be saved at least once ---
    If Len(doc.Path) = 0 Then
        MsgBox "This document has never been saved." & vbCrLf & _
               "Please save it first so a backup can be created.", _
               vbExclamation, "Apply Template Styles"
        Exit Sub
    End If

    ' --- 1. Pick the template from open documents ---
    Set tmplDoc = PickTemplateFromOpenDocs(doc)
    If tmplDoc Is Nothing Then Exit Sub  ' user cancelled

    ' --- 2. Create timestamped backup ---
    Dim backupError As String
    backupPath = CreateBackup(doc, backupError)
    If Len(backupPath) = 0 Then
        MsgBox "Backup creation failed. Aborting to protect your work." & vbCrLf & vbCrLf & _
               "Reason: " & backupError & vbCrLf & _
               "Doc path: " & doc.Path & vbCrLf & _
               "Doc name: " & doc.Name, _
               vbCritical, "Apply Template Styles"
        Exit Sub
    End If

    ' --- 3. Copy headers and footers ---
    currentStep = "Copying headers/footers"
    On Error GoTo ErrHandler
    headersFootersCopied = CopyHeadersFootersFromDoc(doc, tmplDoc)

    currentStep = "Syncing header/footer font"
    headerFooterFontsSynced = SyncHeaderFooterFonts(doc, tmplDoc)

    ' --- 4. Set target document text font from template ---
    currentStep = "Setting text font"
    fontSetFromTemplate = SetBodyFontFromTemplate(doc, tmplDoc, templateFontName)

    ' --- 5. Force target document text to black ---
    currentStep = "Setting text color to black"
    textColorSetToBlack = SetBodyTextColorBlack(doc)

    ' --- 6. Format table headers in target document ---
    currentStep = "Formatting table headers"
    tableHeadersUpdated = FormatTableHeaders(doc, tableHeadersFormatted)

    ' --- 7. Optionally disable auto-update on open ---
    If DISABLE_AUTO_UPDATE Then
        doc.UpdateStylesOnOpen = False
    End If

    ' --- 8. Update fields (page numbers, refs, etc.) ---
    currentStep = "Updating fields"
    doc.Fields.Update

    ' --- 9. Summary ---
    Dim elapsed As Single
    elapsed = Timer - startTime

    Dim summary As String
    summary = "Template formatting applied successfully." & vbCrLf & vbCrLf
    summary = summary & "Template:        " & tmplDoc.Name & vbCrLf
    summary = summary & "Backup:          " & backupPath & vbCrLf
    summary = summary & "Headers/footers: " & IIf(headersFootersCopied, "Copied (exact)", "Skipped (error or no sections)") & vbCrLf
    summary = summary & "H/F style/font:  " & IIf(headerFooterFontsSynced, "Synced from template", "Skipped (error)") & vbCrLf
    If Len(templateFontName) > 0 Then
        summary = summary & "Text font:       " & IIf(fontSetFromTemplate, "Set", "Skipped (error)") & " (""" & templateFontName & """)" & vbCrLf
    Else
        summary = summary & "Text font:       Skipped (template font not found)" & vbCrLf
    End If
    summary = summary & "Text color:      " & IIf(textColorSetToBlack, "Set to black", "Skipped (error)") & vbCrLf

    summary = summary & "Table headers:   " & IIf(tableHeadersUpdated, "Updated", "Skipped (error)") & _
              "; " & tableHeadersFormatted & " of " & doc.Tables.Count & " tables formatted" & vbCrLf

    summary = summary & "Fields updated:  Yes" & vbCrLf
    summary = summary & "Auto-update:     " & IIf(DISABLE_AUTO_UPDATE, "Disabled", "Left unchanged") & vbCrLf
    summary = summary & vbCrLf & "Elapsed: " & Format(elapsed, "0.0") & "s"

    MsgBox summary, vbInformation, "Apply Template Styles - Done"
    Exit Sub

ErrHandler:
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    On Error GoTo 0
    MsgBox "Error " & errNum & ": " & errDesc & vbCrLf & _
           "Step: " & currentStep & vbCrLf & vbCrLf & _
           "Your backup is safe at:" & vbCrLf & backupPath, _
           vbCritical, "Apply Template Styles - Error"
End Sub


' =============================================================================
' Helper: set target text font family to the template's Normal style font.
' =============================================================================
Private Function SetBodyFontFromTemplate(targetDoc As Document, tmplDoc As Document, ByRef outFontName As String) As Boolean

    On Error GoTo FontError

    outFontName = ""

    On Error Resume Next
    outFontName = Trim$(tmplDoc.Styles(wdStyleNormal).Font.Name)
    If Len(outFontName) = 0 Then
        outFontName = Trim$(tmplDoc.Content.Font.Name)
    End If
    On Error GoTo FontError

    If Len(outFontName) = 0 Then
        SetBodyFontFromTemplate = False
        Exit Function
    End If

    targetDoc.Content.Font.Name = outFontName

    Dim storyTypes As Variant
    storyTypes = Array( _
        wdFootnotesStory, _
        wdEndnotesStory, _
        wdCommentsStory, _
        wdTextFrameStory)

    Dim storyType As Variant
    Dim rng As Object

    For Each storyType In storyTypes
        Set rng = Nothing

        On Error Resume Next
        Set rng = targetDoc.StoryRanges(CLng(storyType))
        If Err.Number <> 0 Then
            Err.Clear
            Set rng = Nothing
        End If
        On Error GoTo FontError

        Do While Not rng Is Nothing
            rng.Font.Name = outFontName
            Set rng = rng.NextStoryRange
        Loop
    Next storyType

    SetBodyFontFromTemplate = True
    Exit Function

FontError:
    SetBodyFontFromTemplate = False
End Function


' =============================================================================
' Helper: force body text color to black in main/body-related stories.
' =============================================================================
Private Function SetBodyTextColorBlack(doc As Document) As Boolean

    On Error GoTo ColorError

    doc.Content.Font.Color = wdColorBlack

    Dim storyTypes As Variant
    storyTypes = Array( _
        wdFootnotesStory, _
        wdEndnotesStory, _
        wdCommentsStory, _
        wdTextFrameStory)

    Dim storyType As Variant
    Dim rng As Object

    For Each storyType In storyTypes
        Set rng = Nothing

        On Error Resume Next
        Set rng = doc.StoryRanges(CLng(storyType))
        If Err.Number <> 0 Then
            Err.Clear
            Set rng = Nothing
        End If
        On Error GoTo ColorError

        Do While Not rng Is Nothing
            rng.Font.Color = wdColorBlack
            Set rng = rng.NextStoryRange
        Loop
    Next storyType

    SetBodyTextColorBlack = True
    Exit Function

ColorError:
    SetBodyTextColorBlack = False
End Function


' =============================================================================
' Helper: apply standard header formatting to the first row of every table.
' Matches requested format: gray background + bold text.
' =============================================================================
Private Function FormatTableHeaders(doc As Document, ByRef outTablesFormatted As Long) As Boolean

    On Error GoTo TableHeaderError

    outTablesFormatted = 0

    Dim tbl As Table
    Dim cell As Cell

    For Each tbl In doc.Tables
        On Error Resume Next
        If tbl.Rows.Count > 0 Then
            For Each cell In tbl.Rows(1).Cells
                cell.Shading.BackgroundPatternColor = RGB(191, 191, 191)
                cell.Range.Font.Bold = True
            Next cell
            If Err.Number = 0 Then
                outTablesFormatted = outTablesFormatted + 1
            End If
            Err.Clear
        End If
        On Error GoTo TableHeaderError
    Next tbl

    FormatTableHeaders = True
    Exit Function

TableHeaderError:
    FormatTableHeaders = False
End Function


' =============================================================================
' Helper: Show a dialog listing all open documents (except the active one)
' and let the user pick which one is the template. Returns Nothing if cancelled.
' =============================================================================
Private Function PickTemplateFromOpenDocs(activeDoc As Document) As Document

    ' Build a list of open documents (excluding the active one)
    Dim docNames() As String
    Dim docRefs() As Document
    Dim count As Long
    count = 0

    Dim d As Document
    For Each d In Documents
        If d.FullName <> activeDoc.FullName Then
            count = count + 1
            ReDim Preserve docNames(1 To count)
            ReDim Preserve docRefs(1 To count)
            docNames(count) = d.Name
            Set docRefs(count) = d
        End If
    Next d

    If count = 0 Then
        MsgBox "No other documents are open to use as a template." & vbCrLf & _
               "Please open your template document first.", _
               vbExclamation, "Apply Template Styles"
        Set PickTemplateFromOpenDocs = Nothing
        Exit Function
    End If

    ' If only one other doc is open, confirm it
    If count = 1 Then
        Dim answer As VbMsgBoxResult
        answer = MsgBox("Use """ & docNames(1) & """ as the template?" & vbCrLf & vbCrLf & _
                        "This will copy its headers/footers into:" & vbCrLf & _
                        """" & activeDoc.Name & """", _
                        vbYesNo + vbQuestion, "Apply Template Styles")
        If answer = vbYes Then
            Set PickTemplateFromOpenDocs = docRefs(1)
        Else
            Set PickTemplateFromOpenDocs = Nothing
        End If
        Exit Function
    End If

    ' Multiple docs open — build a numbered list and ask
    Dim prompt As String
    prompt = "Which open document is your template?" & vbCrLf & vbCrLf
    prompt = prompt & "Formatting will be copied INTO: """ & activeDoc.Name & """" & vbCrLf
    prompt = prompt & "Template source document:" & vbCrLf & vbCrLf

    Dim i As Long
    For i = 1 To count
        prompt = prompt & "  " & i & ". " & docNames(i) & vbCrLf
    Next i

    prompt = prompt & vbCrLf & "Enter the number (1-" & count & "):"

    Dim choice As String
    choice = InputBox(prompt, "Apply Template Styles - Select Template")

    If Len(choice) = 0 Then
        Set PickTemplateFromOpenDocs = Nothing
        Exit Function
    End If

    Dim choiceNum As Long
    On Error Resume Next
    choiceNum = CLng(choice)
    On Error GoTo 0

    If choiceNum < 1 Or choiceNum > count Then
        MsgBox "Invalid selection. Please run the macro again.", _
               vbExclamation, "Apply Template Styles"
        Set PickTemplateFromOpenDocs = Nothing
        Exit Function
    End If

    Set PickTemplateFromOpenDocs = docRefs(choiceNum)

End Function


' =============================================================================
' Helper: Create a timestamped backup of the document.
' Returns the backup file path, or "" on failure.
' =============================================================================
Private Function CreateBackup(doc As Document, ByRef outError As String) As String

    On Error GoTo BackupError

    Dim folder As String
    Dim baseName As String
    Dim ext As String
    Dim timestamp As String
    Dim backupName As String

    folder = doc.Path & Application.PathSeparator
    baseName = Left(doc.Name, InStrRev(doc.Name, ".") - 1)
    ext = Mid(doc.Name, InStrRev(doc.Name, "."))
    timestamp = Format(Now, "yyyy-MM-dd_HHmmss")
    backupName = baseName & "_backup_" & timestamp & ext

    Dim backupFullPath As String
    backupFullPath = folder & backupName

    ' Try SaveCopyAs first, fall back to SaveAs2 if unavailable
    On Error Resume Next
    doc.SaveCopyAs backupFullPath
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo BackupError

        ' Fallback: SaveAs2 to the backup path, then re-save original
        Dim originalPath As String
        originalPath = doc.FullName
        doc.SaveAs2 FileName:=backupFullPath
        doc.SaveAs2 FileName:=originalPath
    End If
    On Error GoTo BackupError

    CreateBackup = backupFullPath
    Exit Function

BackupError:
    outError = "Error " & Err.Number & ": " & Err.Description
    CreateBackup = ""
End Function


' =============================================================================
' Helper: Copy headers and footers from an already-open template document.
' Matches sections by index. Returns True on success.
' =============================================================================
Private Function CopyHeadersFootersFromDoc(doc As Document, tmplDoc As Document) As Boolean

    On Error GoTo HFError

    Dim sectionCount As Long
    sectionCount = tmplDoc.Sections.Count
    If sectionCount > doc.Sections.Count Then
        sectionCount = doc.Sections.Count
    End If

    Dim i As Long
    Dim hfType As Variant
    Dim hfTypes As Variant
    hfTypes = Array(wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages)

    For i = 1 To sectionCount
        With doc.Sections(i).PageSetup
            .DifferentFirstPageHeaderFooter = tmplDoc.Sections(i).PageSetup.DifferentFirstPageHeaderFooter
            .OddAndEvenPagesHeaderFooter = tmplDoc.Sections(i).PageSetup.OddAndEvenPagesHeaderFooter
            .HeaderDistance = tmplDoc.Sections(i).PageSetup.HeaderDistance
            .FooterDistance = tmplDoc.Sections(i).PageSetup.FooterDistance
        End With

        For Each hfType In hfTypes
            CopyHeaderFooterExact _
                sourceHF:=tmplDoc.Sections(i).Headers(hfType), _
                targetHF:=doc.Sections(i).Headers(hfType), _
                hasPreviousSection:=(i > 1)

            CopyHeaderFooterExact _
                sourceHF:=tmplDoc.Sections(i).Footers(hfType), _
                targetHF:=doc.Sections(i).Footers(hfType), _
                hasPreviousSection:=(i > 1)
        Next hfType
    Next i

    CopyHeadersFootersFromDoc = True
    Exit Function

HFError:
    CopyHeadersFootersFromDoc = False
End Function


' =============================================================================
' Helper: copy one header/footer story exactly, preserving link behavior.
' =============================================================================
Private Sub CopyHeaderFooterExact( _
    ByVal sourceHF As HeaderFooter, _
    ByVal targetHF As HeaderFooter, _
    ByVal hasPreviousSection As Boolean)

    On Error Resume Next

    ' If template is linked to previous (and previous exists), keep it linked.
    If hasPreviousSection And sourceHF.LinkToPrevious Then
        targetHF.LinkToPrevious = True
        Exit Sub
    End If

    targetHF.LinkToPrevious = False

    If sourceHF.Exists Then
        Dim sourceRng As Range
        Dim targetRng As Range
        Set sourceRng = sourceHF.Range.Duplicate
        Set targetRng = targetHF.Range.Duplicate

        ' FormattedText copies content + formatting without clipboard paste artifacts.
        targetRng.FormattedText = sourceRng.FormattedText
    End If

End Sub


' =============================================================================
' Helper: synchronize header/footer font type from template to target.
' =============================================================================
Private Function SyncHeaderFooterFonts(doc As Document, tmplDoc As Document) As Boolean

    On Error GoTo HFFontError

    Dim headerStyleName As String
    Dim footerStyleName As String

    headerStyleName = tmplDoc.Styles(wdStyleHeader).NameLocal
    footerStyleName = tmplDoc.Styles(wdStyleFooter).NameLocal

    On Error Resume Next
    If Len(headerStyleName) > 0 Then
        Application.OrganizerCopy _
            Source:=tmplDoc.FullName, _
            Destination:=doc.FullName, _
            Name:=headerStyleName, _
            Object:=wdOrganizerObjectStyles
    End If

    If Len(footerStyleName) > 0 Then
        Application.OrganizerCopy _
            Source:=tmplDoc.FullName, _
            Destination:=doc.FullName, _
            Name:=footerStyleName, _
            Object:=wdOrganizerObjectStyles
    End If
    Err.Clear
    On Error GoTo HFFontError

    SyncHeaderFooterFonts = True
    Exit Function

HFFontError:
    SyncHeaderFooterFonts = False
End Function
