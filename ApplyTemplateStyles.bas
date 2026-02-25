' =============================================================================
' ApplyTemplateStyles - Attach Template & Update Styles Macro
' =============================================================================
' Attaches a .dotx/.dotm template to the active document, updates all styles
' to match the template, rebuilds TOC if present, and creates a timestamped
' backup before making any changes.
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

    ' --- 7. Rebuild Table of Contents if present ---
    tocCount = doc.TablesOfContents.Count
    If tocCount > 0 Then
        Dim toc As TableOfContents
        For Each toc In doc.TablesOfContents
            toc.Update
        Next toc
        tocUpdated = True
    End If

    ' --- 8. Update all fields (page numbers, cross-refs, etc.) ---
    doc.Fields.Update

    ' --- 9. Summary ---
    Dim elapsed As Single
    elapsed = Timer - startTime

    Dim summary As String
    summary = "Template styles applied successfully." & vbCrLf & vbCrLf
    summary = summary & "Template:  " & templatePath & vbCrLf
    summary = summary & "Backup:    " & backupPath & vbCrLf
    summary = summary & "Styles in doc:  " & doc.Styles.Count & vbCrLf

    If tocUpdated Then
        summary = summary & "TOC rebuilt:    Yes (" & tocCount & " found)" & vbCrLf
    Else
        summary = summary & "TOC rebuilt:    No TOC found" & vbCrLf
    End If

    summary = summary & "Fields updated: Yes" & vbCrLf
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
        .Title = "Select Your Template (.dotx or .dotm)"
        .Filters.Clear
        .Filters.Add "Word Templates", "*.dotx; *.dotm"
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
