' =============================================================================
' ApplyTemplateStyles - Attach Template & Update Styles Macro
' =============================================================================
' Attaches a .dotx/.dotm template to the active document, updates all styles
' to match the template, copies headers/footers and page setup from the
' template, rebuilds TOC if present, and creates a timestamped backup
' before making any changes.
'
' Usage:
'   1. Open your tech document in Word
'   2. Run the macro (Alt+F8 > ApplyTemplateStyles > Run)
'   3. Select your template file when prompted
'   4. Review the summary dialog
'
' To install:
'   - Open Word > Alt+F11 > Insert > Module > Paste this code
'   - Or save as a .dotm add-in in your Word STARTUP folder
' =============================================================================

Option Explicit

' ---- Configuration ----
' Set this to your template's full path to skip the file picker every time.
' Leave empty ("") to be prompted each run.
Private Const DEFAULT_TEMPLATE_PATH As String = ""

' Set to True to disable "auto update styles on open" after applying.
' Recommended True to prevent unintended changes when opening the doc later.
Private Const DISABLE_AUTO_UPDATE As Boolean = True

Public Sub ApplyTemplateStyles()

    Dim doc As Document
    Dim templatePath As String
    Dim backupPath As String
    Dim styleCountBefore As Long
    Dim tocCount As Long
    Dim tocUpdated As Boolean
    Dim headersFootersCopied As Boolean
    Dim pageSetupCopied As Boolean
    Dim startTime As Single

    ' --- Guard: make sure a document is open ---
    If Documents.Count = 0 Then
        MsgBox "No document is open. Please open a document first.", _
               vbExclamation, "Apply Template Styles"
        Exit Sub
    End If

    Set doc = ActiveDocument
    startTime = Timer

    ' --- Guard: document must be saved at least once (so we know where to put the backup) ---
    If Len(doc.Path) = 0 Then
        MsgBox "This document has never been saved." & vbCrLf & _
               "Please save it first so a backup can be created.", _
               vbExclamation, "Apply Template Styles"
        Exit Sub
    End If

    ' --- 1. Pick the template ---
    templatePath = ResolveTemplatePath()
    If Len(templatePath) = 0 Then Exit Sub  ' user cancelled

    ' Validate the template file exists
    If Dir(templatePath) = "" Then
        MsgBox "Template file not found:" & vbCrLf & templatePath, _
               vbCritical, "Apply Template Styles"
        Exit Sub
    End If

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

    ' --- 3. Count styles before (for summary) ---
    styleCountBefore = doc.Styles.Count

    ' --- 4. Open template ---
    Dim currentStep As String
    currentStep = "Opening template"
    On Error GoTo ErrHandler

    Dim tmplDoc As Document
    Set tmplDoc = Documents.Open( _
        FileName:=templatePath, _
        ReadOnly:=True, _
        AddToRecentFiles:=False, _
        Visible:=False)

    ' --- 5. Copy all styles from template using OrganizerCopy ---
    ' Works with .docx, .dotx, and .dotm — no AttachedTemplate needed
    currentStep = "Copying styles"
    Dim stylesCopied As Long
    Dim s As Style
    For Each s In tmplDoc.Styles
        On Error Resume Next
        Application.OrganizerCopy _
            Source:=templatePath, _
            Destination:=doc.FullName, _
            Name:=s.NameLocal, _
            Object:=wdOrganizerObjectStyles
        If Err.Number = 0 Then
            stylesCopied = stylesCopied + 1
        End If
        Err.Clear
    Next s
    On Error GoTo ErrHandler

    ' --- 6. Copy headers, footers, and page setup directly from open template ---
    currentStep = "Copying headers/footers"
    headersFootersCopied = CopyHeadersFootersFromDoc(doc, tmplDoc)

    currentStep = "Copying page setup"
    pageSetupCopied = CopyPageSetupFromDoc(doc, tmplDoc)

    currentStep = "Closing template"
    tmplDoc.Close SaveChanges:=False
    Set tmplDoc = Nothing

    ' --- 7. Optionally disable auto-update on open ---
    If DISABLE_AUTO_UPDATE Then
        doc.UpdateStylesOnOpen = False
    End If

    ' --- 8. Rebuild Table of Contents if present ---
    tocCount = doc.TablesOfContents.Count
    If tocCount > 0 Then
        Dim toc As TableOfContents
        For Each toc In doc.TablesOfContents
            toc.Update
        Next toc
        tocUpdated = True
    End If

    ' --- 9. Update all fields (page numbers, cross-refs, etc.) ---
    doc.Fields.Update

    ' --- 10. Summary ---
    Dim elapsed As Single
    elapsed = Timer - startTime

    Dim summary As String
    summary = "Template styles applied successfully." & vbCrLf & vbCrLf
    summary = summary & "Template:  " & templatePath & vbCrLf
    summary = summary & "Backup:    " & backupPath & vbCrLf
    summary = summary & "Styles copied:   " & stylesCopied & vbCrLf
    summary = summary & "Styles in doc:   " & doc.Styles.Count & vbCrLf

    summary = summary & "Headers/footers: " & IIf(headersFootersCopied, "Copied from template", "Skipped (error or no sections)") & vbCrLf
    summary = summary & "Page setup:      " & IIf(pageSetupCopied, "Copied from template", "Skipped (error or no sections)") & vbCrLf

    If tocUpdated Then
        summary = summary & "TOC rebuilt:     Yes (" & tocCount & " found)" & vbCrLf
    Else
        summary = summary & "TOC rebuilt:     No TOC found" & vbCrLf
    End If

    summary = summary & "Fields updated:  Yes" & vbCrLf
    summary = summary & "Auto-update on open: " & IIf(DISABLE_AUTO_UPDATE, "Disabled", "Left unchanged") & vbCrLf
    summary = summary & vbCrLf & "Elapsed: " & Format(elapsed, "0.0") & "s"

    MsgBox summary, vbInformation, "Apply Template Styles - Done"
    Exit Sub

ErrHandler:
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If Not tmplDoc Is Nothing Then tmplDoc.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "Error " & errNum & ": " & errDesc & vbCrLf & _
           "Step: " & currentStep & vbCrLf & vbCrLf & _
           "Template: " & templatePath & vbCrLf & _
           "Your backup is safe at:" & vbCrLf & backupPath, _
           vbCritical, "Apply Template Styles - Error"
End Sub


' =============================================================================
' Helper: Resolve the template path (default constant or file picker)
' =============================================================================
Private Function ResolveTemplatePath() As String

    Dim p As String

    ' Use the hardcoded default if set
    If Len(DEFAULT_TEMPLATE_PATH) > 0 Then
        ResolveTemplatePath = DEFAULT_TEMPLATE_PATH
        Exit Function
    End If

    ' Otherwise show a file picker filtered to Word templates
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select Your Template (.docx, .dotx, or .dotm)"
        .Filters.Clear
        .Filters.Add "Word Documents & Templates", "*.docx; *.dotx; *.dotm"
        .Filters.Add "All Files", "*.*"
        .AllowMultiSelect = False

        If .Show = -1 Then
            ' Replace URL-encoded spaces (%20) with actual spaces
            ' FileDialog can return URL-encoded paths that Word can't open
            Dim rawPath As String
            rawPath = .SelectedItems(1)
            ResolveTemplatePath = Replace(rawPath, "%20", " ")
        Else
            ResolveTemplatePath = ""  ' user cancelled
        End If
    End With

End Function


' =============================================================================
' Helper: Create a timestamped backup of the document
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

    ' Try SaveCopyAs first, fall back to SaveAs2 if unavailable
    ' (SaveCopyAs was renamed/removed in some Word versions)
    Dim backupFullPath As String
    backupFullPath = folder & backupName

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
' Matches sections by index — if the doc has more sections than the template,
' extra sections keep their existing headers/footers. Returns True on success.
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
        ' Match the "Different First Page" and "Different Odd & Even" settings
        doc.Sections(i).PageSetup.DifferentFirstPageHeaderFooter = _
            tmplDoc.Sections(i).PageSetup.DifferentFirstPageHeaderFooter
        doc.Sections(i).PageSetup.OddAndEvenPagesHeaderFooter = _
            tmplDoc.Sections(i).PageSetup.OddAndEvenPagesHeaderFooter

        For Each hfType In hfTypes
            ' Copy headers
            If tmplDoc.Sections(i).Headers(hfType).Exists Then
                tmplDoc.Sections(i).Headers(hfType).Range.Copy
                doc.Sections(i).Headers(hfType).Range.Paste

                ' Preserve "Link to Previous" setting
                doc.Sections(i).Headers(hfType).LinkToPrevious = _
                    tmplDoc.Sections(i).Headers(hfType).LinkToPrevious
            End If

            ' Copy footers
            If tmplDoc.Sections(i).Footers(hfType).Exists Then
                tmplDoc.Sections(i).Footers(hfType).Range.Copy
                doc.Sections(i).Footers(hfType).Range.Paste

                ' Preserve "Link to Previous" setting
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
' Transfers margins, orientation, paper size, gutter, section start type,
' vertical alignment, and header/footer distances. Matches by section index.
' Returns True on success.
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
            ' Margins
            .TopMargin = tmplDoc.Sections(i).PageSetup.TopMargin
            .BottomMargin = tmplDoc.Sections(i).PageSetup.BottomMargin
            .LeftMargin = tmplDoc.Sections(i).PageSetup.LeftMargin
            .RightMargin = tmplDoc.Sections(i).PageSetup.RightMargin
            .Gutter = tmplDoc.Sections(i).PageSetup.Gutter
            .GutterPos = tmplDoc.Sections(i).PageSetup.GutterPos

            ' Orientation and paper size
            .Orientation = tmplDoc.Sections(i).PageSetup.Orientation
            .PaperSize = tmplDoc.Sections(i).PageSetup.PaperSize
            .PageWidth = tmplDoc.Sections(i).PageSetup.PageWidth
            .PageHeight = tmplDoc.Sections(i).PageSetup.PageHeight

            ' Header and footer distances from edge
            .HeaderDistance = tmplDoc.Sections(i).PageSetup.HeaderDistance
            .FooterDistance = tmplDoc.Sections(i).PageSetup.FooterDistance

            ' Section start type (new page, continuous, even page, odd page)
            .SectionStart = tmplDoc.Sections(i).PageSetup.SectionStart

            ' Vertical alignment (top, center, justified, bottom)
            .VerticalAlignment = tmplDoc.Sections(i).PageSetup.VerticalAlignment

            ' Mirror margins (for bound documents)
            .MirrorMargins = tmplDoc.Sections(i).PageSetup.MirrorMargins
        End With
    Next i

    CopyPageSetupFromDoc = True
    Exit Function

PSError:
    CopyPageSetupFromDoc = False
End Function
