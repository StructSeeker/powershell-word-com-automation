<#
.SYNOPSIS
    Initialize PowerShell REPL with persistent Word connection
.DESCRIPTION
    Connects to active Word instance and loads helper functions
    for interactive document automation. Provides convenient functions
    for testing and exploring Word COM operations in a persistent REPL session.

    CONNECTION STRATEGY:
    Uses [Microsoft.VisualBasic.Interaction]::GetObject(path) to connect to the
    specific Word instance that has the target document open. This avoids the
    common pitfall of GetActiveObject/GetObject(progID) returning a headless
    background Word process instead of the visible UI instance.
.NOTES
    Session logs to: .logs/word-repl-YYYYMMDD.log
    Global variables: $word, $doc (Word and Document COM objects)
    Example: pwsh -NoExit -Command ". 'word-repl-init.ps1'"
            (dot-source — using & spawns a child scope and functions are lost)
#>

# Configure logging (uses working directory for portability)
$LogPath = Join-Path (Get-Location) ".logs"
$LogFile = Join-Path $LogPath "word-repl-$(Get-Date -Format 'yyyyMMdd').log"

# Ensure log directory exists
if (-not (Test-Path $LogPath)) {
    $null = New-Item -ItemType Directory -Path $LogPath -Force
}

# ── Force UTF-8 console encoding (critical for powershell 5.1) ──
# Without this, Write-Host output of non-ASCII characters (bullets, checkmarks,
# box-drawing) is garbled when the script is invoked from powershell.exe.
# pwsh uses UTF-8 by default and is unaffected, but these calls are harmless there.
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Logging function
function Write-REPLLog {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Severity = 'INFO'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Severity] $Message"
    
    try {
        Add-Content -Path $LogFile -Value $logEntry -Encoding UTF8 -ErrorAction Stop
    } catch {
        # Silent fail—don't break REPL if logging fails
    }
}

try {
    Write-Host "Initializing Word REPL session..." -ForegroundColor Cyan
    Write-REPLLog "REPL session started (PS $($PSVersionTable.PSVersion))"
    
    # ── Load Microsoft.VisualBasic for GetObject ──
    # GetActiveObject (Marshal) doesn't exist in pwsh/.NET Core, and even in
    # powershell it often returns a headless Word background process rather
    # than the visible UI instance. GetObject(path) is the only reliable way
    # to reach the correct Word process.
    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
    
    # ── Discover the correct document ──
    # Strategy: scan for .docx files in the working directory, try GetObject
    # on each until we find one that's actually open in Word.
    $candidates = @(Get-ChildItem -Path (Get-Location) -Filter '*.docx' | Select-Object -ExpandProperty FullName)
    $candidates += @(Get-ChildItem -Path (Get-Location) -Filter '*.doc' | Select-Object -ExpandProperty FullName)
    
    $global:doc = $null
    $global:word = $null
    
    # First try Paper.docx specifically (most common case)
    $preferred = @(Join-Path (Get-Location) 'Paper.docx')
    $preferred += $candidates | Where-Object { $_ -like '*Paper*' }
    $preferred += $candidates
    
    $tried = @()
    foreach ($candidate in ($preferred | Select-Object -Unique)) {
        if ($global:doc -ne $null) { break }
        if (-not (Test-Path $candidate)) { continue }
        $tried += $candidate
        try {
            $global:doc = [Microsoft.VisualBasic.Interaction]::GetObject($candidate, '')
            $global:word = $global:doc.Application
            Write-REPLLog "Connected via GetObject(path): $candidate"
        } catch {
            # Document not open in any Word instance—skip
            continue
        }
    }
    
    if ($null -eq $global:doc) {
        # Fallback: try GetObject(progID) for headless instances
        Write-Host "  ⚠ No open .docx found via GetObject(path). Trying fallback..." -ForegroundColor Yellow
        try {
            $global:word = [Microsoft.VisualBasic.Interaction]::GetObject('', 'Word.Application')
            if ($global:word.Documents.Count -gt 0) {
                $global:doc = $global:word.ActiveDocument
                Write-REPLLog "Connected via GetObject(progID): $($global:doc.Name)"
            } else {
                throw "GetObject(progID) returned Word with 0 open documents (likely a headless instance)."
            }
        } catch {
            throw "Cannot connect to a Word instance with an open document. Ensure Word is running and Paper.docx (or another .docx) is open.`nTried paths: $($tried -join '; ')`nFallback error: $_"
        }
    }
    
    Write-Host "✓ Connected to Word document: $($global:doc.Name)" -ForegroundColor Green
    Write-Host "✓ Document has $($global:doc.Paragraphs.Count) paragraphs, $($global:doc.Comments.Count) comments" -ForegroundColor Green
    Write-REPLLog "Connected to: $($global:doc.Name)"
    
} catch {
    $errorMsg = "Failed to connect to Word: $_"
    Write-Host "❌ $errorMsg" -ForegroundColor Red
    Write-REPLLog $errorMsg -Severity 'ERROR'
    Write-Host ""
    Write-Host "Troubleshooting:" -ForegroundColor Yellow
    Write-Host "  1. Is Word running?" -ForegroundColor Gray
    Write-Host "  2. Is Paper.docx (or another .docx) open in Word?" -ForegroundColor Gray
    Write-Host "  3. Close hidden headless Word processes: Get-Process WINWORD | Where-Object { !$_.MainWindowTitle } | Stop-Process" -ForegroundColor Gray
    exit 1
}

# ==============================================================================
# REPL Helper Functions
# ==============================================================================

<#
.SYNOPSIS
    Display document information
.OUTPUTS
    Hashtable with document metadata (Document, Path, Paragraphs, Comments, Saved)
.EXAMPLE
    PS> info
#>
function info {
    $docInfo = @{
        Document = $global:doc.Name
        Path = $global:doc.FullName
        Paragraphs = $global:doc.Paragraphs.Count
        Comments = $global:doc.Comments.Count
        Saved = $global:doc.Saved
    }
    
    Write-REPLLog "info: Retrieved document info"
    return $docInfo
}

<#
.SYNOPSIS
    Add comment to paragraph
.PARAMETER ParagraphIndex
    1-based paragraph index
.PARAMETER Text
    Comment text
.OUTPUTS
    Comment COM object
.EXAMPLE
    PS> $comment = add-comment -ParagraphIndex 5 -Text "Review this"
#>
function add-comment {
    param(
        [int]$ParagraphIndex,
        [string]$Text
    )
    
    try {
        $para = $global:doc.Paragraphs.Item($ParagraphIndex)
        $comment = $global:doc.Comments.Add($para.Range, $Text)
        Write-Host "✓ Comment added to paragraph $ParagraphIndex" -ForegroundColor Green
        Write-REPLLog "Comment added: para $ParagraphIndex - $($Text.Substring(0, [Math]::Min(50, $Text.Length)))..."
        return $comment
    } catch {
        Write-Error "Failed to add comment to paragraph $ParagraphIndex`: $_"
        Write-REPLLog "Failed to add comment: $_" -Severity 'ERROR'
    }
}

<#
.SYNOPSIS
    Get paragraph text and COM object
.PARAMETER Index
    1-based paragraph index
.OUTPUTS
    Hashtable with Index, Text, and ComObject
.EXAMPLE
    PS> $para = get-paragraph -Index 5
    PS> $para.Text
#>
function get-paragraph {
    param([int]$Index)
    
    try {
        $para = $global:doc.Paragraphs.Item($Index)
        $text = $para.Range.Text
        
        Write-Host "Paragraph $Index`:" -ForegroundColor Cyan
        Write-Host $text -ForegroundColor Yellow
        
        Write-REPLLog "Retrieved paragraph $Index"
        
        return @{
            Index    = $Index
            Text     = $text
            ComObject = $para
        }
    } catch {
        Write-Error "Failed to retrieve paragraph $Index`: $_"
        Write-REPLLog "Failed to get paragraph $Index`: $_" -Severity 'ERROR'
    }
}

<#
.SYNOPSIS
    List all comments in document
.OUTPUTS
    Array of hashtables with comment information
.EXAMPLE
    PS> get-comments
#>
function get-comments {
    $count = $global:doc.Comments.Count
    Write-Host "Total comments: $count`n" -ForegroundColor Cyan
    
    $commentsList = @()
    
    for ($i = 1; $i -le $count; $i++) {
        $comment = $global:doc.Comments.Item($i)
        $text = $comment.Range.Text
        
        Write-Host "[$i] $($comment.Author): $($text.Substring(0, [Math]::Min(70, $text.Length)))..." -ForegroundColor Yellow
        
        $commentsList += @{
            Index     = $i
            Author    = $comment.Author
            Text      = $text
            ComObject = $comment
        }
    }
    
    Write-REPLLog "Listed $count comments"
    return $commentsList
}

<#
.SYNOPSIS
    Delete all comments from document
.OUTPUTS
    Int (number of deleted comments)
.EXAMPLE
    PS> remove-comments
#>
function remove-comments {
    try {
        $count = $global:doc.Comments.Count
        
        # Iterate backwards to avoid index shifting
        for ($i = $count; $i -ge 1; $i--) {
            $global:doc.Comments.Item($i).Delete()
        }
        
        Write-Host "✓ Deleted $count comments" -ForegroundColor Green
        Write-REPLLog "Deleted $count comments"
        
        return $count
    } catch {
        Write-Error "Failed to delete comments: $_"
        Write-REPLLog "Failed to delete comments: $_" -Severity 'ERROR'
    }
}

<#
.SYNOPSIS
    Save document
.PARAMETER Path
    Optional new file path
.EXAMPLE
    PS> save-doc
    PS> save-doc -Path "C:\new-file.docx"
#>
function save-doc {
    param([string]$Path)
    
    try {
        if ([string]::IsNullOrEmpty($Path)) {
            $global:doc.Save()
            Write-Host "✓ Document saved" -ForegroundColor Green
            Write-REPLLog "Document saved"
        } else {
            $global:doc.SaveAs2($Path)
            Write-Host "✓ Document saved as: $Path" -ForegroundColor Green
            Write-REPLLog "Document saved as: $Path"
        }
    } catch {
        Write-Error "Failed to save document: $_"
        Write-REPLLog "Failed to save document: $_" -Severity 'ERROR'
    }
}

# ==============================================================================
# Session Info
# ==============================================================================

Write-Host "`n📝 Word REPL Session Active" -ForegroundColor Green
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Gray
Write-Host "Available helper functions:" -ForegroundColor Cyan
Write-Host ""
Write-Host "  info                                - Show document metadata" -ForegroundColor Gray
Write-Host "  add-comment -ParagraphIndex 5 -Text 'text'  - Add comment" -ForegroundColor Gray
Write-Host "  get-paragraph -Index 5              - Get paragraph text & object" -ForegroundColor Gray
Write-Host "  get-comments                        - List all comments" -ForegroundColor Gray
Write-Host "  remove-comments                     - Delete all comments" -ForegroundColor Gray
Write-Host "  save-doc [-Path 'path']             - Save document" -ForegroundColor Gray
Write-Host ""
Write-Host "Direct COM access:" -ForegroundColor Cyan
Write-Host "  `$word                              - Word application COM object" -ForegroundColor Gray
Write-Host "  `$doc                               - Document COM object" -ForegroundColor Gray
Write-Host "  `$doc.Paragraphs.Count              - Direct property access" -ForegroundColor Gray
Write-Host ""
Write-Host "Session log: $LogFile" -ForegroundColor Gray
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Gray
Write-Host ""

