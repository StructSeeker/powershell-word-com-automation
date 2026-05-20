---
name: powershell-word-com-automation
description: Automates Microsoft Word via COM interop using PowerShell (pwsh/powershell). Covers connection via GetObject(path), orphan process cleanup, WdBuiltInStyle integers, encoding setup, content building, REPL exploration, and production scripting.
when_to_use: |
  - User asks to automate Microsoft Word document editing or generation
  - User needs to write or modify PowerShell scripts that interact with Word documents
  - User encounters Word COM errors like GetActiveObject not found, 0 documents, garbled text, or locale-dependent style failures
  - User needs to manipulate .docx files programmatically
allowed-tools:
  - Read
  - Write
  - Grep
  - RunInTerminal
shell: powershell
paths:
  - "**/*.ps1"
  - "**/*.docx"
effort: high
tags:
  - powershell
  - COM automation
  - Microsoft Word
  - scripting
user-invocable: true
---

# PowerShell Word COM Automation

**Purpose:** Automate Microsoft Word via COM interop with zero dependencies. Supports both interactive REPL sessions and production scripts.

**Key constraint:** Scripts always operate on files in the **current working directory** — never inside the skill directory.

---

## Pre-flight Checklist

Before writing any automation, run through this in order:

1. **Kill orphan Word processes** — headless `WINWORD.exe` instances from prior sessions will intercept `GetObject` calls
2. **Verify the document path** — use an absolute path resolved from `Get-Location`
3. **Confirm encoding** — save scripts as UTF-8 with BOM; set console encoding at the top of every script
4. **Connect via `GetObject(absolute-path)`** — this is the only reliable connection method

---

## Quick Reference

| Task | Command |
|------|---------|
| **Start REPL** | `. .\path\to\word-repl-init.ps1` (dot-source) |
| **Run script** | `pwsh -ExecutionPolicy Bypass -File script.ps1` |
| **Dot-source in existing terminal** | `. .\script.ps1` |
| **Connect to Word** | `Add-Type -AssemblyName Microsoft.VisualBasic` then `GetObject(absolute-path)` — see below |
| **Kill orphans** | `Get-Process WINWORD \| Where-Object { -not $_.MainWindowTitle } \| Stop-Process -Force` |

---

## Two Shells, One Rule

| Shell | .NET runtime | Default encoding | Invoke with |
|-------|-------------|-----------------|-------------|
| **`pwsh`** (PowerShell 7+) | .NET Core / .NET 5+ | UTF-8 | `pwsh -File script.ps1` |
| **`powershell`** (5.1) | .NET Framework | System-dependent | `powershell -File script.ps1` |

**Always use `-File`**, never `-Command "& 'script.ps1'"` — the latter mangles UTF-8 output.

**Prefer dot-sourcing** (`. .\script.ps1`) in an existing terminal over spawning a new `pwsh` process. It keeps COM objects in scope for interactive inspection and avoids unnecessary process overhead.

---

## Connection (Critical)

### The only reliable method

```powershell
Add-Type -AssemblyName Microsoft.VisualBasic
$doc = [Microsoft.VisualBasic.Interaction]::GetObject("C:\full\path\to\Paper.docx", "")
$word = $doc.Application
```

`GetObject(absolute-path)` follows the file to whichever Word process owns it, even if multiple Word instances exist.

### Why other methods fail

| Method | Problem |
|--------|---------|
| `[Marshal]::GetActiveObject("Word.Application")` | Does not exist in pwsh (.NET Core) |
| `[VB.Interaction]::GetObject("", "Word.Application")` | Returns a headless background Word instance (no visible window, 0 documents) |

### Kill orphans first

Repeated automation sessions leave headless `WINWORD.exe` processes that silently intercept `GetObject`:

```powershell
Get-Process WINWORD -ErrorAction SilentlyContinue |
    Where-Object { -not $_.MainWindowTitle } | Stop-Process -Force
Start-Sleep -Milliseconds 500
```

---

## Encoding

### Rules (apply to every script)

```powershell
# At the top of every script
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
$PSDefaultParameterValues['Add-Content:Encoding'] = 'utf8'
```

- **Save all `.ps1` files as UTF-8 with BOM.** (VS Code: *File > Save with Encoding > UTF-8 with BOM*)
- **Pass `-Encoding UTF8`** on every explicit file write
- **Use ASCII-safe characters** in string literals: `--` not `—`, `"` not `""`, `...` not `…`

The console encoding lines are critical for `powershell.exe` 5.1 compatibility; pwsh 7 handles this by default but they are harmless to include.

---

## Styles: WdBuiltInStyle Integers (Not String Names)

Style names are locale-dependent — `"Heading 1"` fails on Chinese or Spanish Word. Always use integer constants instead.

```powershell
# Correct — locale-independent
$word.Selection.Style = $doc.Styles.Item(-2)   # wdStyleHeading1
$word.Selection.Style = $doc.Styles.Item(-3)   # wdStyleHeading2
$word.Selection.Style = $doc.Styles.Item(-4)   # wdStyleHeading3
$word.Selection.Style = $doc.Styles.Item(-1)   # wdStyleNormal

# Wrong — fails on localized Word
$word.Selection.Style = "Heading 1"
```

| Style | WdBuiltInStyle value |
|-------|---------------------|
| `wdStyleNormal` | `-1` |
| `wdStyleHeading1` | `-2` |
| `wdStyleHeading2` | `-3` |
| `wdStyleHeading3` | `-4` |

---

## Building / Replacing Document Content

### The right way to clear a document

```powershell
# Correct — preserves paragraph structure
$doc.Range().Delete()
Start-Sleep -Milliseconds 200

# Wrong — corrupts the paragraph model; causes Paragraphs.Add() to produce malformed content
$doc.Content.Delete()
```

### The right way to insert content

After clearing, use `$word.Selection.TypeText()` + `TypeParagraph()`. Do **not** use `$doc.Paragraphs.Add()` after a full content wipe — it produces malformed paragraphs.

```powershell
$doc.Range().Delete()
Start-Sleep -Milliseconds 200

$word.Selection.Style = $doc.Styles.Item(-2)    # wdStyleHeading1
$word.Selection.TypeText("My Title") | Out-Null
$word.Selection.TypeParagraph() | Out-Null

$word.Selection.Style = $doc.Styles.Item(-1)    # wdStyleNormal
$word.Selection.TypeText("Body text here.") | Out-Null
$word.Selection.TypeParagraph() | Out-Null
```

### Helper functions

These cover the vast majority of content-building needs. Accept `$doc` and `$word` as explicit parameters rather than relying on globals — makes them portable and testable.

```powershell
function Add-Heading {
    param(
        [string]$Text,
        [int]$Level = 1,
        $Word,
        $Doc
    )
    $styleIndex = -2 - ($Level - 1)  # Level 1 -> -2, Level 2 -> -3, Level 3 -> -4
    $Word.Selection.Style = $Doc.Styles.Item($styleIndex)
    $Word.Selection.TypeText($Text) | Out-Null
    $Word.Selection.TypeParagraph() | Out-Null
    Start-Sleep -Milliseconds 80
}

function Add-Paragraph {
    param(
        [string]$Text,
        $Word,
        $Doc
    )
    $Word.Selection.Style = $Doc.Styles.Item(-1)   # wdStyleNormal
    $Word.Selection.TypeText($Text) | Out-Null
    $Word.Selection.TypeParagraph() | Out-Null
    Start-Sleep -Milliseconds 40
}
```

---

## Script Mode

### Minimal production skeleton

```powershell
# ── Encoding (required for powershell 5.1 compatibility) ──
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
$PSDefaultParameterValues['Add-Content:Encoding'] = 'utf8'

try {
    Add-Type -AssemblyName Microsoft.VisualBasic

    # 1. Kill orphan Word processes
    Get-Process WINWORD -ErrorAction SilentlyContinue |
        Where-Object { -not $_.MainWindowTitle } | Stop-Process -Force
    Start-Sleep -Milliseconds 500

    # 2. Connect via document path (absolute)
    $docPath = Join-Path (Get-Location) "Paper.docx"
    $doc  = [Microsoft.VisualBasic.Interaction]::GetObject($docPath, "")
    $word = $doc.Application

    # 3. Clear content (Range.Delete — never Content.Delete)
    $doc.Range().Delete()
    Start-Sleep -Milliseconds 200

    # 4. Write content using Selection + WdBuiltInStyle integers
    $word.Selection.Style = $doc.Styles.Item(-2)
    $word.Selection.TypeText("Title Here") | Out-Null
    $word.Selection.TypeParagraph() | Out-Null

    # --- your operations here ---

    # 5. Save
    $doc.Save()
    Write-Host "Done"

} catch {
    Write-Error "Failed: $_"
    exit 1
}
```

### Targeting a specific open document by name

```powershell
$docName = "Paper.docx"
$doc = $null
Get-ChildItem (Get-Location) -Filter '*.docx' | ForEach-Object {
    try {
        $d = [Microsoft.VisualBasic.Interaction]::GetObject($_.FullName, "")
        if ($d.Name -eq $docName) { $doc = $d }
    } catch { }  # not open in Word — skip
}
if (-not $doc) { throw "'$docName' not found or not open in Word" }
$word = $doc.Application
```

---

## REPL Mode (Interactive Exploration)

```powershell
# From your project root (where Paper.docx lives)
# IMPORTANT: dot-source with (.) — using & spawns a child scope and
# helper functions will not survive into your interactive session
. 'path\to\word-repl-init.ps1'

# Or launch pwsh and stay interactive:
pwsh -NoExit -Command ". 'path\to\word-repl-init.ps1'"
```

### Available commands once connected

| Command | Purpose |
|---------|---------|
| `info` | Show document metadata |
| `get-paragraph -Index 5` | Get paragraph text |
| `add-comment -ParagraphIndex 5 -Text "..."` | Add a comment |
| `get-comments` | List all comments |
| `remove-comments` | Delete all comments |
| `save-doc [-Path '...']` | Save document |

### Typical session

```powershell
info
get-paragraph -Index 3
add-comment -ParagraphIndex 3 -Text "Expand this section"
get-comments
save-doc
```

You can dot-source helper scripts from within the REPL to load utility functions without leaving the session:
```powershell
. .\add-review-comments.ps1
Add-ReviewComments -Comments @(
    @{ ParagraphIndex = 5; Text = "Strong opening" }
)
```

---

## COM Object Lifetime

**Do not call `Marshal.ReleaseComObject`** in Word automation scripts. While the method exists in pwsh 7 on Windows (`[SupportedOSPlatform("windows")]`), the official .NET guidance warns that improper use can cause access violations and subtle process corruption — risks that are compounded because `Word.Application` is effectively a singleton. The .NET GC releases COM references correctly when objects fall out of scope. Just let it do its job.

If a Word process does not exit after your script completes, the cause is almost always a headless orphan from a prior session — not an unreleased RCW. Use the orphan cleanup pattern above.

---

## Common COM Operations

| Task | Code |
|------|------|
| Get document name | `$doc.Name` |
| Get paragraph count | `$doc.Paragraphs.Count` |
| Get paragraph text | `$doc.Paragraphs.Item($i).Range.Text` |
| Get selection text | `$word.Selection.Text` |
| Add comment | `$doc.Comments.Add($range, "text")` |
| Format text bold | `$range.Font.Bold = $true` |
| Set font size | `$range.Font.Size = 14` |
| Save | `$doc.Save()` |
| Save as | `$doc.SaveAs2("path.docx")` |

**⚠ All COM indexing is 1-based.** `.Item(1)` is the first item.

**⚠ Add `Start-Sleep -Milliseconds 100–150` between rapid operations** to give COM time to process.

---

## Working Directory Discipline

The skill directory is never the working directory. All scripts and logs go to `$pwd` (the project root).

```powershell
# Correct
$LogFile = Join-Path (Get-Location) ".logs\word-automation-$(Get-Date -Format 'yyyyMMdd').log"
$docPath = Join-Path (Get-Location) "Paper.docx"

# Wrong — don't write into the skill folder
$LogFile = Join-Path $PSScriptRoot ".logs\session.log"
```

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---------|-------------|-----|
| `GetActiveObject` not found | Running pwsh (.NET Core) | Use `GetObject(path)` instead |
| Connected but 0 documents | Hit a headless Word instance | Kill orphans, reconnect via `GetObject(path)` |
| Document has no content after connecting | Connected to wrong instance | Ensure correct absolute path in `GetObject` |
| Garbled text in logs / console | Encoding mismatch | Add encoding block at top; save script as UTF-8 with BOM |
| Script blocked | Execution policy | Use `-ExecutionPolicy Bypass` |
| Operations are flaky or out of order | Rapid COM calls | Insert `Start-Sleep` between operations |
| `Style = "Heading 1"` fails | Localized Word install | Use `$doc.Styles.Item(-2)` integer constant |
| `Paragraphs.Add()` produces malformed content | `Content.Delete()` used to clear | Use `Range.Delete()` then `Selection.TypeText()` |
| Word process lingers after script | Usually a pre-existing headless orphan | Run orphan cleanup; not a `ReleaseComObject` issue |

---

## DO / DON'T

**DO:**
- Connect via `GetObject(absolute-path)` — always
- Kill headless Word processes before connecting
- Use `$doc.Range().Delete()` to clear content
- Use `$word.Selection.TypeText()` + `TypeParagraph()` to insert content
- Use `WdBuiltInStyle` integers (`Styles.Item(-2)`, etc.)
- Wrap operations in `try/catch`
- Add small delays between COM calls
- Save scripts as UTF-8 with BOM
- Write logs to `$pwd\.logs\`
- Dot-source scripts (`. .\script.ps1`) in existing terminals when possible
- Pass `$doc` and `$word` explicitly into helper functions

**DON'T:**
- Use `GetActiveObject` or `GetObject("", "Word.Application")`
- Use `$doc.Content.Delete()` — corrupts the paragraph model
- Use `$doc.Paragraphs.Add()` after a full document wipe
- Use string style names like `"Heading 1"` — locale-dependent
- Use `Marshal.ReleaseComObject` — dangerous for singleton COM objects; GC handles it
- Write files into the skill directory
- Use smart quotes (`""`), em-dashes (`—`), or non-ASCII in string literals
- Omit `-ExecutionPolicy Bypass` when invoking scripts
- Use `-Command "& 'script.ps1'"` — use `-File` instead