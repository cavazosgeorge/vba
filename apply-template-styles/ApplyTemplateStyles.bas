' =============================================================================
' ApplyTemplateStyles - Copy Styles, Headers/Footers & Page Setup from Template
' =============================================================================
' Copies all styles, headers/footers, and page layout from a template document
' into the active document. Creates a timestamped backup before making changes.
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
' Set to True to remove manual font/paragraph overrides so style definitions win.
Private Const CLEAR_DIRECT_FORMATTING As Boolean = True

Public Sub ApplyTemplateStyles()

    Dim doc As Document
    Dim tmplDoc As Document
    Dim backupPath As String
    Dim stylesCopied As Long
    Dim tocCount As Long
    Dim tocUpdated As Boolean
    Dim headersFootersCopied As Boolean
    Dim pageSetupCopied As Boolean
    Dim directFormattingCleared As Boolean
    Dim tablesFormatted As Long
    Dim tableStyleName As String
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

    ' --- 3. Copy all styles from template using OrganizerCopy ---
    currentStep = "Copying styles"
    On Error GoTo ErrHandler

    Dim tmplPath As String
    Dim docPath As String
    tmplPath = tmplDoc.FullName
    docPath = doc.FullName

    Dim s As Style
    For Each s In tmplDoc.Styles
        On Error Resume Next
        Application.OrganizerCopy _
            Source:=tmplPath, _
            Destination:=docPath, _
            Name:=s.NameLocal, _
            Object:=wdOrganizerObjectStyles
        If Err.Number = 0 Then
            stylesCopied = stylesCopied + 1
        End If
        Err.Clear
    Next s
    On Error GoTo ErrHandler

    ' --- 4. Clear direct text formatting so copied styles can take effect ---
    ' This removes manual font/paragraph overrides while preserving layout-level
    ' formatting such as table borders, section breaks, and headers/footers.
    currentStep = "Clearing direct formatting"
    If CLEAR_DIRECT_FORMATTING Then
        directFormattingCleared = ClearDirectFormatting(doc)
    End If

    ' --- 5. Copy headers, footers, and page setup ---
    currentStep = "Copying headers/footers"
    headersFootersCopied = CopyHeadersFootersFromDoc(doc, tmplDoc)

    currentStep = "Copying page setup"
    pageSetupCopied = CopyPageSetupFromDoc(doc, tmplDoc)

    ' --- 6. Apply template's table style to all tables in target doc ---
    currentStep = "Formatting tables"
    If tmplDoc.Tables.Count > 0 Then
        ' Use first template table as style/look source.
        Dim tmplTable As Table
        Set tmplTable = tmplDoc.Tables(1)
        tableStyleName = CStr(tmplTable.Style)

        Dim useHeadingRows As Boolean
        Dim useLastRow As Boolean
        Dim useFirstCol As Boolean
        Dim useLastCol As Boolean
        Dim useRowBands As Boolean
        Dim useColBands As Boolean
        useHeadingRows = tmplTable.ApplyStyleHeadingRows
        useLastRow = tmplTable.ApplyStyleLastRow
        useFirstCol = tmplTable.ApplyStyleFirstColumn
        useLastCol = tmplTable.ApplyStyleLastColumn
        useRowBands = tmplTable.ApplyStyleRowBands
        useColBands = tmplTable.ApplyStyleColumnBands

        ' Apply style and style options to every table in the target.
        Dim tbl As Table
        For Each tbl In doc.Tables
            On Error Resume Next
            tbl.Style = tableStyleName
            tbl.ApplyStyleHeadingRows = useHeadingRows
            tbl.ApplyStyleLastRow = useLastRow
            tbl.ApplyStyleFirstColumn = useFirstCol
            tbl.ApplyStyleLastColumn = useLastCol
            tbl.ApplyStyleRowBands = useRowBands
            tbl.ApplyStyleColumnBands = useColBands
            If Err.Number = 0 Then
                tablesFormatted = tablesFormatted + 1
            End If
            Err.Clear
        Next tbl
        On Error GoTo ErrHandler
    End If

    ' --- 7. Optionally disable auto-update on open ---
    If DISABLE_AUTO_UPDATE Then
        doc.UpdateStylesOnOpen = False
    End If

    ' --- 8. Rebuild Table of Contents if present ---
    currentStep = "Rebuilding TOC"
    tocCount = doc.TablesOfContents.Count
    If tocCount > 0 Then
        Dim toc As TableOfContents
        For Each toc In doc.TablesOfContents
            toc.Update
        Next toc
        tocUpdated = True
    End If

    ' --- 9. Update all fields (page numbers, cross-refs, etc.) ---
    currentStep = "Updating fields"
    doc.Fields.Update

    ' --- 10. Summary ---
    Dim elapsed As Single
    elapsed = Timer - startTime

    Dim summary As String
    summary = "Template styles applied successfully." & vbCrLf & vbCrLf
    summary = summary & "Template:        " & tmplDoc.Name & vbCrLf
    summary = summary & "Backup:          " & backupPath & vbCrLf
    summary = summary & "Styles copied:   " & stylesCopied & vbCrLf
    summary = summary & "Styles in doc:   " & doc.Styles.Count & vbCrLf
    summary = summary & "Headers/footers: " & IIf(headersFootersCopied, "Copied", "Skipped (error or no sections)") & vbCrLf
    summary = summary & "Page setup:      " & IIf(pageSetupCopied, "Copied", "Skipped (error or no sections)") & vbCrLf
    If CLEAR_DIRECT_FORMATTING Then
        summary = summary & "Direct format:   " & IIf(directFormattingCleared, "Cleared (text stories)", "Skipped (error)") & vbCrLf
    Else
        summary = summary & "Direct format:   Left unchanged" & vbCrLf
    End If

    If Len(tableStyleName) > 0 Then
        summary = summary & "Tables:          " & tablesFormatted & " of " & doc.Tables.Count & " formatted as """ & tableStyleName & """" & vbCrLf
    Else
        summary = summary & "Tables:          No tables found in template" & vbCrLf
    End If

    If tocUpdated Then
        summary = summary & "TOC rebuilt:     Yes (" & tocCount & " found)" & vbCrLf
    Else
        summary = summary & "TOC rebuilt:     No TOC found" & vbCrLf
    End If

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
                        "This will copy its styles, headers/footers, and page setup into:" & vbCrLf & _
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
    prompt = prompt & "Styles will be copied INTO: """ & activeDoc.Name & """" & vbCrLf
    prompt = prompt & "Styles will be copied FROM the document you choose:" & vbCrLf & vbCrLf

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
        doc.Sections(i).PageSetup.DifferentFirstPageHeaderFooter = _
            tmplDoc.Sections(i).PageSetup.DifferentFirstPageHeaderFooter
        doc.Sections(i).PageSetup.OddAndEvenPagesHeaderFooter = _
            tmplDoc.Sections(i).PageSetup.OddAndEvenPagesHeaderFooter

        For Each hfType In hfTypes
            If tmplDoc.Sections(i).Headers(hfType).Exists Then
                tmplDoc.Sections(i).Headers(hfType).Range.Copy
                doc.Sections(i).Headers(hfType).Range.Paste
                doc.Sections(i).Headers(hfType).LinkToPrevious = _
                    tmplDoc.Sections(i).Headers(hfType).LinkToPrevious
            End If

            If tmplDoc.Sections(i).Footers(hfType).Exists Then
                tmplDoc.Sections(i).Footers(hfType).Range.Copy
                doc.Sections(i).Footers(hfType).Range.Paste
                doc.Sections(i).Footers(hfType).LinkToPrevious = _
                    tmplDoc.Sections(i).Footers(hfType).LinkToPrevious
            End If
        Next hfType
    Next i

    CopyHeadersFootersFromDoc = True
    Exit Function

HFError:
    CopyHeadersFootersFromDoc = False
End Function


' =============================================================================
' Helper: Copy page setup properties from an already-open template document.
' Matches by section index. Returns True on success.
' =============================================================================
Private Function CopyPageSetupFromDoc(doc As Document, tmplDoc As Document) As Boolean

    On Error GoTo PSError

    Dim sectionCount As Long
    sectionCount = tmplDoc.Sections.Count
    If sectionCount > doc.Sections.Count Then
        sectionCount = doc.Sections.Count
    End If

    Dim i As Long
    For i = 1 To sectionCount
        With doc.Sections(i).PageSetup
            .TopMargin = tmplDoc.Sections(i).PageSetup.TopMargin
            .BottomMargin = tmplDoc.Sections(i).PageSetup.BottomMargin
            .LeftMargin = tmplDoc.Sections(i).PageSetup.LeftMargin
            .RightMargin = tmplDoc.Sections(i).PageSetup.RightMargin
            .Gutter = tmplDoc.Sections(i).PageSetup.Gutter
            .GutterPos = tmplDoc.Sections(i).PageSetup.GutterPos
            .Orientation = tmplDoc.Sections(i).PageSetup.Orientation
            .PaperSize = tmplDoc.Sections(i).PageSetup.PaperSize
            .PageWidth = tmplDoc.Sections(i).PageSetup.PageWidth
            .PageHeight = tmplDoc.Sections(i).PageSetup.PageHeight
            .HeaderDistance = tmplDoc.Sections(i).PageSetup.HeaderDistance
            .FooterDistance = tmplDoc.Sections(i).PageSetup.FooterDistance
            .SectionStart = tmplDoc.Sections(i).PageSetup.SectionStart
            .VerticalAlignment = tmplDoc.Sections(i).PageSetup.VerticalAlignment
            .MirrorMargins = tmplDoc.Sections(i).PageSetup.MirrorMargins
        End With
    Next i

    CopyPageSetupFromDoc = True
    Exit Function

PSError:
    CopyPageSetupFromDoc = False
End Function


' =============================================================================
' Helper: remove direct character/paragraph overrides from text stories only.
' This lets style definitions drive appearance without stripping layout objects.
' =============================================================================
Private Function ClearDirectFormatting(doc As Object) As Boolean

    On Error GoTo ClearError

    Dim storyTypes As Variant
    storyTypes = Array( _
        wdMainTextStory, _
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
        On Error GoTo ClearError

        If Not rng Is Nothing Then
            ClearDirectFormattingInStoryChain rng
        End If
    Next storyType

    ClearDirectFormatting = True
    Exit Function

ClearError:
    ClearDirectFormatting = False
End Function


' =============================================================================
' Helper: each story type can have linked ranges. Clear each linked range.
' =============================================================================
Private Sub ClearDirectFormattingInStoryChain(ByVal storyRng As Object)

    Dim rng As Object
    Dim para As Object
    Set rng = storyRng

    Do While Not rng Is Nothing
        ' Word version compatibility:
        ' Some builds vary in available Range members.
        ' Reset direct formatting paragraph-by-paragraph and skip table cells
        ' so table borders/shading are preserved for the table-style step.
        For Each para In rng.Paragraphs
            On Error Resume Next
            If para.Range.Tables.Count = 0 Then
                para.Range.Font.Reset
                para.Range.ParagraphFormat.Reset
            End If
            On Error GoTo 0
        Next para

        Set rng = rng.NextStoryRange
    Loop

End Sub
