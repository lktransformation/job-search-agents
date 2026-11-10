# CV and Cover Letter Agent

A Python script that drafts a tailored CV and a one-page cover letter for a specific job, using **your own** career history and style rules. It researches the company, checks your experience against the job spec, flags gaps and weak spots, and only produces the Word documents after you approve the draft.

It uses Claude through the Claude Agent SDK. If you want to use ChatGPT, Gemini or another AI instead, use the copy-and-paste prompt in [`../prompts/cv-cover-letter-agent.md`](../prompts/cv-cover-letter-agent.md) (see [Use any AI chat](../docs/04-use-any-ai-chat.md)).

> **This is a drafting tool.** Check every claim, figure and date before you send anything.

## Before you start

You need to have done these (see the main [README](../README.md)):

1. Installed Python, Claude Code and the Python packages: [install guide](../docs/01-install-prerequisites.md)
2. Logged in to Claude Code by running `claude` once
3. **Created your own files** (below)

## Step 1: Create your files

The repo does not include any career data. You must create these three files yourself from the templates.

Open a terminal in this folder (`cv-cover-letter-agent`) and run:

**Windows (PowerShell):**

```powershell
Copy-Item profile.example.json profile.json
Copy-Item knowledge\cv_style_guide.example.md knowledge\cv_style_guide.md
Copy-Item knowledge\cover_letter_style_guide.example.md knowledge\cover_letter_style_guide.md
```

**Mac or Linux:**

```bash
cp profile.example.json profile.json
cp knowledge/cv_style_guide.example.md knowledge/cv_style_guide.md
cp knowledge/cover_letter_style_guide.example.md knowledge/cover_letter_style_guide.md
```

| File | What to put in it |
|---|---|
| `profile.json` | Your name, phone, email and location line. Replace John Smith's details. See the field guide below. |
| `knowledge/cv_style_guide.md` | Your career history (the only facts the agent may use), your tools, your qualifications and your CV rules. Build the career history with AI using the [guide](../docs/02-build-your-career-history.md). |
| `knowledge/cover_letter_style_guide.md` | Your tone, sentence style, punctuation rules and cover letter structure. Review it using the [voice and style guide](../docs/03-set-your-voice-and-style.md). |

**When you have finished editing each style guide, delete `(TEMPLATE)` from its first line.** The script will refuse to run while it is still there, so you cannot forget to fill them in.

### profile.json fields

| Field | Meaning |
|---|---|
| `full_name` | Printed at the top of the CV and cover letter, and used for the sign-off |
| `phone`, `email` | Printed in the contact line |
| `location_line` | For example `United Kingdom (Remote / Hybrid)` |
| `qualification_line` | Optional. A certification to show at the end of the contact line. Leave it empty if you have none. |
| `experience_framing_note` | Optional but recommended. One or two honest sentences about how much experience you have in your target specialism versus general experience, so the agent never overstates it. |
| `output_prefix` | Used in the file names, for example `John_Smith` |

## Step 2: Run the agent

In the terminal, still inside the `cv-cover-letter-agent` folder:

**Windows:**

```powershell
.venv\Scripts\python tailor_application.py
```

**Mac or Linux:**

```bash
.venv/bin/python tailor_application.py
```

These use the private Python environment from step 5 of the [install guide](../docs/01-install-prerequisites.md). If `.venv` does not exist yet, do that step first.

The script asks you three things:

1. **Company website**, for example `https://www.examplecompany.com`
2. **Role title**, for example `Senior Programme Manager`
3. **Job spec.** Paste the whole text, then type `END` on a new line and press Enter.

## Step 3: Review the draft

The agent will research the company, check your experience against the job spec, and show you a review with three kinds of flags:

- **Gap flags:** requirements your career history does not cover
- **Quality flags:** bullets missing a "how" or a measurable result
- **Naturalness flags:** cover letter links that feel stretched

Answer its questions and give it any missing detail. It will not paper over a real gap. If your experience does not match the core requirements, it will tell you plainly and may recommend not applying.

## Step 4: Approve and collect your files

When you are happy, type `approved` (or "looks good" or "go ahead"). The script saves two Word documents into an `output` folder:

```
output/CompanyName_RoleTitle/John_Smith_CV_CompanyName.docx
output/CompanyName_RoleTitle/John_Smith_Cover_Letter_CompanyName.docx
```

Open them, read every line, and check every figure before you send them.

To stop without generating files, type `quit`.

## Troubleshooting

- **"profile.json not found":** run the copy commands in Step 1.
- **"still contains the example details":** open `profile.json` and replace the example email and name with your own.
- **"still marked (TEMPLATE)":** finish editing the style guide, then delete `(TEMPLATE)` from its first line.
- **Login or authentication error:** run `claude` once and sign in. See the [Claude Agent SDK documentation](https://platform.claude.com/docs/en/agent-sdk/python) if it persists.
- **`ModuleNotFoundError`, or `.venv` not found:** run the package install commands in step 5 of the [install guide](../docs/01-install-prerequisites.md) again.
