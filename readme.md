# Word Document Automation

This guide explains how to work with your AI agent to automate Microsoft Word documents. The automation runs through PowerShell COM interop — the agent writes and executes scripts that talk directly to a live Word instance on your machine. No add-ins, no macros, no dependencies.

---

## How It Works (The Mental Model)

The agent connects to Word by finding an **already-open document** on your filesystem. It does not open files itself. This means the setup is always:

1. **You** open the document in Word
2. **You** tell the agent which file to work on and where it lives
3. **The agent** connects to it, performs the operations, and saves

Think of it like handing the agent the keys to a car that's already running — you still need to start it.

---

## Before You Ask the Agent to Do Anything

### 1. Open your document in Word

Just double-click it normally. Word must be running with the file open and visible. If Word shows the document in a background/protected view banner, click **Enable Editing** first — COM cannot write to protected-view documents.

### 2. Know your file's full path

An absolute path is required to connect reliably. The easiest ways to get it:

- In Windows Explorer, hold **Shift** and right-click the file → **Copy as path**
- In Word's title bar, hover over the document name — the full path appears as a tooltip
- In PowerShell: `(Get-Item ".\MyDoc.docx").FullName`

A path looks like: `C:\Users\YourName\Documents\Project\Report.docx`

### 3. Close other documents you don't want touched

If you have multiple `.docx` files open, tell the agent explicitly which one to work on (see below). Having only one document open eliminates ambiguity entirely.

---

## Telling the Agent Which Document to Use

Always be explicit. These phrasings work well:

> "The file is at `C:\Users\Alice\Documents\Thesis\Chapter1.docx`"

> "I have `Report.docx` open in Word. It's in `C:\Work\Q3\`"

> "There are two docs open — work on `Draft_v2.docx`, not `Draft_v1.docx`"

**Why this matters:** The agent connects via the file path. If you have multiple documents open and don't specify, the agent may connect to the wrong one. There is no undo for COM operations — changes are written directly to the live document.

---

## Choosing a Mode: REPL vs Script

There are two ways the agent can work. Choose based on what you need.

### REPL Mode — for exploration and step-by-step work

REPL (Read-Eval-Print Loop) is an interactive PowerShell session where the agent can send commands one at a time and see results between each step. Use this when:

- You want to inspect the document first (paragraph count, existing content, comments)
- You're not sure exactly what you want yet and want to iterate
- You want to review each change before proceeding
- You're doing something exploratory like reviewing and annotating a draft

**How to ask for it:**

> "Start a REPL session on my document so we can explore it together"

> "Open an interactive session — I want to check the structure before making changes"

In REPL mode, the agent will ask you to run commands in a terminal and paste back any output it needs to see.

### Script Mode — for well-defined, repeatable operations

Script mode generates a complete `.ps1` file that runs start to finish. Use this when:

- You know exactly what you want done
- The task is large (rewriting sections, reformatting the whole document, batch commenting)
- You want something you can re-run later
- You don't need to inspect the document first

**How to ask for it:**

> "Write a script to replace the entire document with this new content: ..."

> "Create a script that adds review comments to every paragraph that mentions 'Q3'"

> "Generate a script I can save and re-run whenever I update this template"

**When in doubt, REPL first.** It's safer — you can stop at any point. Script mode is faster for large or well-understood tasks.

---

## Multiple Documents Open at the Same Time

If you have more than one `.docx` open, always specify the target by name **and** path. Just a filename is not always enough if the same name exists in different folders.

**Good:**
> "Work on `Proposal.docx` — it's at `C:\Projects\ClientA\Proposal.docx`"

**Risky:**
> "Work on the proposal" *(the agent will ask, but saves a round-trip if you pre-empt it)*

**If you want the agent to list what's open first:**
> "List all the Word documents currently open before doing anything"

The agent can enumerate open documents and confirm with you before touching anything.

---

## What to Include in Your Request

The more context you give upfront, the less back-and-forth. A complete request looks like:

> "I have `Chapter3.docx` open in Word at `C:\Thesis\Chapter3.docx`. I want you to:
> 1. Add a comment on paragraph 4 saying 'Needs a citation'
> 2. Add a comment on paragraph 11 saying 'Expand this argument'
> Use script mode — I'll run it myself."

For content replacement:

> "Replace the entire content of `Template.docx` (`C:\Work\Templates\Template.docx`) with the following structure: [paste content here]. Use Heading 1 for section titles, normal style for body text."

---

## Running Scripts the Agent Gives You

When the agent produces a `.ps1` script, run it like this:

```powershell
pwsh -ExecutionPolicy Bypass -File .\script.ps1
```

Or if you already have a PowerShell terminal open in the right folder, dot-source it:

```powershell
. .\script.ps1
```

**Always read the script before running it.** The agent will explain what each section does, but you're the one executing it — and COM operations write directly to your live document.

---

## Before Each Session: Quick Safety Checklist

| Step | Why |
|------|-----|
| Document is open in Word | COM cannot connect to a closed file |
| **Enable Editing** is clicked (if prompted) | Protected view blocks all writes |
| You know the full file path | Required for reliable connection |
| You've told the agent which file if multiple are open | Prevents operating on the wrong document |
| You've saved a backup if the changes are destructive | COM writes are immediate; there is no undo from the agent's side |

---

## Common Scenarios

### "Add comments to a document for review"

Best in **script mode** if you know which paragraphs to annotate. Best in **REPL** if you want to read paragraphs first and decide as you go.

> "I want to review `Draft.docx` at `C:\Writing\Draft.docx`. Start a REPL so I can read each paragraph and tell you which ones to comment on."

### "Rewrite / replace the whole document"

Use **script mode**. Give the agent the complete desired content in your message.

> "Replace everything in `Report.docx` (`C:\Reports\Report.docx`) with the following content. Keep it in script mode so I can review the script first."

### "Check what's in the document before changing anything"

Use **REPL mode** and ask the agent to call `info` and list some paragraphs first.

> "Before making any changes, start a REPL on `Notes.docx` at `C:\Notes\Notes.docx` and show me the paragraph count and the first 5 paragraphs."

### "I have a template I want to fill out repeatedly"

Use **script mode** and ask the agent to generate a reusable script with parameters.

> "Create a parameterized script for `Invoice.docx` at `C:\Templates\Invoice.docx` that I can call with a client name and amount each time."

---

## Things The Agent Cannot Do

- **Open a file for you** — you must open it in Word first
- **Undo changes** — COM writes directly to the live document; save a backup beforehand if needed
- **Work on a cloud-only file** — the document must be locally available (OneDrive files need to be synced/downloaded first)
- **Work on password-protected documents** — COM cannot authenticate
- **See your screen** — The agent works blind; describe what you're seeing if something goes wrong

---

## If Something Goes Wrong

**"Connected but the document looks empty / wrong content"**
The agent may have connected to the wrong Word instance. Close any extra documents, kill orphan Word processes (the agent can do this), and try again with the explicit file path.

**"The script ran but nothing changed"**
Check that the document wasn't in Protected View. Click Enable Editing and re-run.

**"Word froze or became unresponsive"**
A COM operation hung. Force-close Word via Task Manager, reopen your document (saved changes should be there if `$doc.Save()` ran), and ask the agent to retry with smaller steps.

**"I see garbled characters in the output"**
Encoding mismatch. Tell the agent you're on Windows PowerShell 5.1 (not pwsh 7) if that's what you're using — it will add the appropriate encoding headers.

---

## Tips for Smooth Sessions

- **Save your document before starting** — gives you a clean restore point
- **Close unrelated documents** — removes ambiguity entirely
- **Paste the full file path directly** — saves at least one round-trip every session
- **For large tasks, describe the structure** — e.g. "the document has an intro, three sections with subsections, and a conclusion" helps the agent make smarter decisions
- **Prefer REPL for first-time tasks** — once you've done something once and it worked, script mode is faster for repeats
