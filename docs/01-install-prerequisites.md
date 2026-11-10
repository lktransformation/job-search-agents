# Install the prerequisites

This guide assumes you have never used a terminal. It takes about 15 minutes.

**Using ChatGPT, Gemini, Copilot or another chat tool instead?** Skip this guide and go to [Use any AI chat](04-use-any-ai-chat.md).

## What you need

| Tool | Needed for | Required? |
|---|---|---|
| Python 3.10 or newer | The CV and cover letter agent | Yes, for that agent |
| Claude Code | The interview prep agent, and the CV agent's login | Yes, for both agents |
| A Claude account on a paid plan (Pro, Max, Team, Enterprise) or an Anthropic Console account | Claude Code. The free Claude plan does not include it. | Yes |
| Python packages `claude-agent-sdk` and `python-docx` (listed in `cv-cover-letter-agent/requirements.txt`) | The CV and cover letter agent | Yes, for that agent |
| Node.js | Only if you want the optional Playwright browser tool for the interview prep agent | Optional |
| Git for Windows | Optional extra for Claude Code on Windows | Optional |

## 1. Open a terminal

A terminal is a window where you type commands.

- **Windows:** press the Windows key, type `PowerShell`, and press Enter. Your prompt starts with `PS C:\`.
- **Mac:** press Cmd + Space, type `Terminal`, and press Enter.

**Tip for Windows:** to open a terminal already inside a folder, open the folder in File Explorer, click the address bar at the top, type `powershell`, and press Enter.

To move into a folder, use `cd` followed by the folder path. For example, if you unzipped the repo into Documents:

```powershell
cd $HOME\Documents\job-search-agents
```

```bash
cd ~/Documents/job-search-agents
```

Type `dir` (Windows) or `ls` (Mac) to list what is in the current folder.

## 2. Install Python

**Windows (easiest):** in PowerShell, run:

```powershell
winget install Python.Python.3.12
```

Close PowerShell and open a new window. Or download the installer from [python.org/downloads](https://www.python.org/downloads/) and **tick "Add python.exe to PATH"** on the first screen of the installer.

**Mac:** download the installer from [python.org/downloads](https://www.python.org/downloads/) and run it, or if you use Homebrew run `brew install python`.

**Check it worked:**

```powershell
python --version
```

On Mac, use `python3 --version`. You should see `Python 3.10` or higher. On Windows, if `python` is not recognised, try `py --version`.

## 3. Install Claude Code

Claude Code needs a paid Claude plan. Run the command for your system, then open a **new** terminal window.

**Windows PowerShell:**

```powershell
irm https://claude.ai/install.ps1 | iex
```

**Mac, Linux, or WSL:**

```bash
curl -fsSL https://claude.ai/install.sh | bash
```

**Windows alternative (WinGet):**

```powershell
winget install Anthropic.ClaudeCode
```

**Check it worked:**

```
claude --version
```

You should see a version number. If your terminal says `claude` is not found, close it, open a new one and try again. If it still fails, see the [Claude Code troubleshooting guide](https://code.claude.com/docs/en/troubleshoot-install).

## 4. Log in to Claude Code

```
claude
```

Follow the browser prompts to sign in with your Claude account. Once you see the Claude Code prompt, type `/exit` to leave for now.

## 5. Install the Python packages (CV and cover letter agent)

The packages go into a private folder called `.venv` inside `cv-cover-letter-agent`, so they cannot clash with anything else on your computer. Installing them system-wide with a plain `pip install` fails on many Macs and Linux systems with an `externally-managed-environment` error, so use these commands instead. From the repo folder, run:

**Windows (PowerShell):**

```powershell
cd cv-cover-letter-agent
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
cd ..
```

If `python` is not recognised, use `py -m venv .venv` for the second line.

**Mac or Linux:**

```bash
cd cv-cover-letter-agent
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
cd ..
```

On Linux, if the second line says `ensurepip is not available`, install your distribution's `python3-venv` package and run it again.

The `claude-agent-sdk` package includes what it needs to run, so you do not have to install anything else for it. If the script ever reports a login or authentication error, see the [Claude Agent SDK documentation](https://platform.claude.com/docs/en/agent-sdk/python).

## 6. Optional: Node.js for the Playwright browser tool

The interview prep agent can optionally use a Playwright browser tool to open job posting pages that normal fetching cannot read. This needs Node.js. Download the LTS version from [nodejs.org](https://nodejs.org/en/download), run the installer, then check:

```
node --version
```

You can skip this and paste job specs in as text instead.

## You are ready

Go back to the main [README](../README.md) and continue from Step 3.

## If something goes wrong

- **"command not found" or "not recognised":** close the terminal and open a new one. Installers only update your system path for new windows.
- **PowerShell shows `'irm' is not recognized`:** you are in Command Prompt, not PowerShell. Open PowerShell instead.
- **`externally-managed-environment` or permission errors from pip:** you ran `pip` directly. Use the `.venv` commands in step 5 instead.
- **Anything else:** paste the exact error message into any AI chat and ask it to explain what to do. Tell it your operating system.
