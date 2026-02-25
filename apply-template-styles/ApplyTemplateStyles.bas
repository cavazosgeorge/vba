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
    backupPath = CreateBackup(doc)
    If Len(backupPath) = 0 Then
        MsgBox "Backup creation failed. Aborting to protect your work.", _
               vbCritical, "Apply Template Styles"
        Exit Sub
    End If

    ' --- 3. Count styles before (for summary) ---
    styleCountBefore = doc.Styles.Count

    ' --- 4. Attach the template ---
    On Error GoTo ErrHandler
    doc.AttachedTemplate = templatePath

    ' --- 5. Update all styles from the template ---
    doc.UpdateStyles

    ' --- 6. Optionally disable auto-update on open ---
    If DISABLE_AUTO_UPDATE Then
        doc.UpdateStylesOnOpen = False
    End If

    ' --- 7. Copy headers, footers, and page setup from template ---
    headersFootersCopied = CopyHeadersFooters(doc, templatePath)
    pageSetupCopied = CopyPageSetup(doc, templatePath)

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
    summary = summary & "Styles in doc:  " & doc.Styles.Count & vbCrLf

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
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
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
            ResolveTemplatePath = .SelectedItems(1)
        Else
            ResolveTemplatePath = ""  ' user cancelled
        End If
    End With

End Function


' =============================================================================
' Helper: Create a timestamped backup of the document
' Returns the backup file path, or "" on failure.
' =============================================================================
Private Function CreateBackup(doc As Document) As String

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

    ' Save current state first so the backup reflects latest edits
    If Not doc.Saved Then
        doc.Save
    End If

    ' Copy the file
    FileCopy folder & doc.Name, folder & backupName

    CreateBackup = folder & backupName
    Exit Function

BackupError:
    CreateBackup = ""
End Function


' =============================================================================
' Helper: Copy headers and footers from the template into the active document.
' Opens the template as a hidden document, copies header/footer content from
' each section, then closes it. Matches sections by index — if the doc has
' more sections than the template, extra sections keep their existing
' headers/footers. Returns True on success.
' =============================================================================
Private Function CopyHeadersFooters(doc As Document, templatePath As String) As Boolean

    On Error GoTo HFError

    Dim tmplDoc As Document
    Set tmplDoc = Documents.Open( _
        FileName:=templatePath, _
        ReadOnly:=True, _
        AddToRecentFiles:=False, _
        Visible:=False)

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

    tmplDoc.Close SaveChanges:=False
    CopyHeadersFooters = True
    Exit Function

HFError:
    On Error Resume Next
    If Not tmplDoc Is Nothing Then tmplDoc.Close SaveChanges:=False
    CopyHeadersFooters = False
End Function


' =============================================================================
' Helper: Copy page setup properties from the template into the active document.
' Transfers margins, orientation, paper size, gutter, section start type,
' vertical alignment, and header/footer distances. Matches by section index.
' Returns True on success.
' =============================================================================
Private Function CopyPageSetup(doc As Document, templatePath As String) As Boolean

    On Error GoTo PSError

    Dim tmplDoc As Document
    Set tmplDoc = Documents.Open( _
        FileName:=templatePath, _
        ReadOnly:=True, _
        AddToRecentFiles:=False, _
        Visible:=False)

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

    tmplDoc.Close SaveChanges:=False
    CopyPageSetup = True
    Exit Function

PSError:
    On Error Resume Next
    If Not tmplDoc Is Nothing Then tmplDoc.Close SaveChanges:=False
    CopyPageSetup = False
End Function
