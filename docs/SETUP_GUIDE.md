# Setup guide: Claude + Power BI MCP in VS Code

This guide gets a Windows machine from nothing to a working connection between **Claude Code** in **VS Code** and **Power BI Desktop**, through Microsoft's Power BI MCP server. Part A is for the IT team; Part B is for each participant.

*Checked on 19 November 2026 against winget, the VS Code Marketplace, Anthropic's Claude Code docs and Microsoft's Power BI docs. Anything marked **(unverified)** could not be confirmed from an official source and may look slightly different on your machine.*

---

# Part A: For the IT team

## At a glance

Every participant machine needs:

- [ ] **Power BI Desktop**, latest version
- [ ] **VS Code** with the **Claude Code** extension
- [ ] **Node.js 22 or newer** (includes `npm` and `npx`)
- [ ] **Git for Windows**
- [ ] **Claude Code CLI** (the `claude` command)
- [ ] A **paid Claude plan for each user**: Pro, Max, Team or Enterprise. The Free plan does **not** include Claude Code.
- [ ] **Local admin rights** during setup
- [ ] **Internet access** to the hosts in A3
- [ ] **PowerShell scripts allowed** for the user (see A3)
- [ ] **SQL Server Developer or Express + SSMS**: not needed yet, needed from about Session 7

## A1. Software

| Software | Version | Notes | Install command (winget) |
| --- | --- | --- | --- |
| Power BI Desktop | Latest | Runs the report and its local analysis engine | `winget install --id Microsoft.PowerBI -e` |
| Visual Studio Code | Latest | Where participants work with Claude | `winget install --id Microsoft.VisualStudioCode -e` |
| Claude Code extension for VS Code | Latest | Publisher: Anthropic | `code --install-extension anthropic.claude-code` |
| Node.js (LTS) | 22 or newer | Provides `npx`, which downloads and runs the Power BI MCP server | `winget install --id OpenJS.NodeJS.LTS -e` |
| Git for Windows | Latest | Needed from Session 2 (Git diff). Also recommended by Anthropic so Claude Code can use its Bash tool | `winget install --id Git.Git -e` |
| Claude Code CLI | Latest | Installs **per user, no admin needed**. Run as the participant, not as admin | `irm https://claude.ai/install.ps1 \| iex` (PowerShell) |
| SQL Server 2025 Developer *or* Express | Latest | From about Session 7. Developer is free for training and development use | `winget install --id Microsoft.SQLServer.2025.Developer -e` or `winget install --id Microsoft.SQLServer.2025.Express -e` |
| SQL Server Management Studio 22 | Latest | From about Session 7 | `winget install --id Microsoft.SQLServerManagementStudio.22 -e` |

Notes:

- All winget IDs above were checked with `winget show --id <id> --exact`. The plain `Microsoft.SQLServerManagementStudio` ID installs an **older** version (20.x); use `Microsoft.SQLServerManagementStudio.22`.
- Claude Code alternatives: `winget install --id Anthropic.ClaudeCode -e` (winget installs **don't auto-update**; upgrade with `winget upgrade Anthropic.ClaudeCode`), or from CMD: `curl -fsSL https://claude.ai/install.cmd -o install.cmd && install.cmd && del install.cmd`.
- Claude Code needs **Windows 10 1809+** or Windows Server 2019+.
- The VS Code extension does **not** put the `claude` command in the terminal. Install the CLI as well.

## A2. Accounts and licences

- **One paid Claude plan per user**: Pro, Max, Team or Enterprise (Anthropic also lists Ultra and Console accounts). The free claude.ai plan does not include Claude Code.
- **Sign-in happens in the browser** (claude.ai). The **Claude Desktop app is not needed**.
- **Local admin rights** are needed while installing Power BI Desktop, Node.js, Git and SQL Server. The Claude Code CLI itself installs per user.
- **Power BI Desktop features:** participants must be able to save as a **Power BI Project (.pbip)** with the **TMDL** format. In older Desktop versions these are switches under *Preview features* (see Part B step 3); please don't block them by policy.

## A3. Network and security

**Outbound HTTPS (port 443)** must be allowed to the hosts below. The list comes from Anthropic's [network configuration page](https://code.claude.com/docs/en/network-config); check it for changes before rollout, because it is the source of truth.

| Host | Needed for |
| --- | --- |
| `api.anthropic.com` | Claude Code itself |
| `claude.ai`, `claude.com`, `platform.claude.com`, `sso.claude.com` | Signing in |
| `downloads.claude.ai` | The Claude Code installer and auto-updates |
| `registry.npmjs.org` | Downloading the Power BI MCP server with `npx` |
| `github.com`, `raw.githubusercontent.com` | Plugin marketplaces, release notes; this training's materials |
| `storage.googleapis.com` | Plugin metadata (and installer versions before 2.1.116) |
| `code.claude.com` | Documentation lookups |
| `mcp-proxy.anthropic.com` | claude.ai connectors (optional for this training) |

Also:

- **Proxies:** set `HTTPS_PROXY` (and `NO_PROXY`) for the user, written with a scheme, e.g. `http://proxy.example.com:8080`. SOCKS proxies are not supported.
- **SSL inspection:** the inspection certificate must be trusted. Claude Code trusts the Windows certificate store by default; otherwise point `NODE_EXTRA_CA_CERTS` at the CA file (`.pem`). **Node.js and npx need the same certificate**, or the MCP server download fails.
- **PowerShell execution policy:** allow the user to run local scripts:

  ```
  Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
  ```

- **Don't block localhost.** Power BI Desktop runs a local analysis engine on a random `localhost` port, and the MCP server connects to it there.
- **Data policy:** use **sample data only** in this training. Microsoft's own warning for the MCP server: AI assistance *"may transmit accessed data to the LLM provider"*. Query results go into the Claude conversation. No client or personal data.

## A4. Verification (run on each machine, as the participant)

Open a **new** PowerShell window and run:

| Command | Healthy result |
| --- | --- |
| `node --version` | `v22` or higher |
| `git --version` | Any version |
| `code --version` | A version number |
| `claude --version` | A version number and `(Claude Code)` |
| `claude doctor` | No errors |
| `code --list-extensions` | Includes `anthropic.claude-code` |

## A5. Frequent issues

| Symptom | Fix |
| --- | --- |
| `'claude' is not recognized` | The terminal was open during install. Close it and open a new one. If it persists, sign out of Windows and back in, or check the user's PATH |
| `running scripts is disabled on this system` | Run `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`, then open a new terminal |
| `node --version` shows v20 or lower | Upgrade: `winget upgrade --id OpenJS.NodeJS.LTS -e` |
| Sign-in page won't open, or loops | Set a default browser. Behind a proxy, check that `claude.ai`, `platform.claude.com` and `sso.claude.com` are allowed. In the terminal you can press `c` to copy the sign-in link and paste it into a browser |
| `npx` can't download the server (timeouts, `SELF_SIGNED_CERT_IN_CHAIN`, `UNABLE_TO_GET_ISSUER_CERT`) | `registry.npmjs.org` is blocked, or the SSL-inspection certificate isn't trusted by Node. Allow the host and set `NODE_EXTRA_CA_CERTS` |
| Sign-in works but Claude Code says there's no access | The user is on the Free plan. Upgrade to a paid plan |
| No option to save as `.pbip` | Update Power BI Desktop, then check *Preview features* (Part B step 3) |
| Claude can't find, or picks the wrong, Power BI Desktop | More than one Desktop window is open. Close all but one |

---

# Part B: For participants

Do these steps in order. Each one ends with a **Check**: don't move on until it passes.

## Step 1. Requirements check

Open VS Code, then a terminal (**Ctrl + \`**). Make sure the terminal says **PowerShell**. Type:

```
node --version
git --version
code --version
```

**Check:** Node shows `v22` or higher, and the other two show a version number. If anything says *not recognized*, ask your IT team (Part A).

## Step 2. One-time Git setup

Tell Git your name and email (used to label your changes):

```
git config --global user.name "Your Name"
git config --global user.email "you@example.com"
```

**Check:** `git config --global user.name` prints your name.

## Step 3. Power BI Desktop settings

1. Open Power BI Desktop, then **File › Options and settings › Options › Preview features**.
2. If you see **Power BI Project (.pbip) save option** and **Store semantic model using TMDL format**, tick both. *(unverified: Microsoft's current docs no longer call these preview features, so in newer versions they may already be on and not listed.)*
3. Restart Power BI Desktop.

**Check:** **File › Save as** offers **Power BI project files (\*.pbip)**.

## Step 4. Install Claude Code

In the **PowerShell** terminal:

```
irm https://claude.ai/install.ps1 | iex
```

If your terminal is **CMD** instead:

```
curl -fsSL https://claude.ai/install.cmd -o install.cmd && install.cmd && del install.cmd
```

Then **close the terminal and open a new one** (the new one finds the `claude` command), and run:

```
claude --version
claude doctor
```

**Check:** `claude --version` shows a version, and `claude doctor` reports no errors.

## Step 5. Sign in to Claude

You need a paid Claude plan (Pro, Max, Team or Enterprise). No Claude Desktop app needed.

**Option A: in the terminal**

1. Type `claude` and press Enter. (Already inside Claude? Type `/login`.)
2. Choose the option for your **Claude account with a subscription** *(unverified: the exact wording of this option isn't in Anthropic's docs; pick the claude.ai / subscription option, not the "Anthropic Console account" one, unless your company uses Console)*.
3. Your browser opens on claude.ai. Sign in (email link or code, Google, or your company's single sign-on).
4. Click **Authorize**. If the terminal asks for a code, paste it.
5. Type `/exit` to leave Claude.

**Option B: in the VS Code extension**

1. Install the extension: Extensions (**Ctrl + Shift + X**), search **Claude Code** (publisher **Anthropic**), Install. Or in the terminal: `code --install-extension anthropic.claude-code`.
2. Open the Claude Code panel and click **Sign in**.
3. Sign in in the browser and click **Authorize**. You're sent back to VS Code.

The extension does **not** add the `claude` command to your terminal. You still need Step 4 for Step 6.

**Check:** the terminal shows **Login successful**, or the VS Code panel lets you type a message.

## Step 6. Register the Power BI MCP server (read-only, once)

In **PowerShell**, type this as one line:

```
claude mcp add --scope user powerbi-modeling '--' cmd /c npx -y @microsoft/powerbi-modeling-mcp@latest --start --readonly
```

What the parts mean:

- `claude mcp add` registers a server with Claude Code.
- `--scope user` registers it **once for your whole account**, so it works in every folder. Without it, it only works in the folder you're in.
- `powerbi-modeling` is the name we give it.
- `'--'` separates Claude's options from the server's command. The **quotes matter in PowerShell**: without them, PowerShell can swallow the two dashes and you get `error: unknown option '-y'`.
- `cmd /c npx -y @microsoft/powerbi-modeling-mcp@latest --start` downloads and starts Microsoft's server.
- `--readonly` is Microsoft's **safe mode**: it *"prevents any write operations to your semantic model"*. Claude can read, not change.

In **CMD**, use the same command **without** quotes around the dashes:

```
claude mcp add --scope user powerbi-modeling -- cmd /c npx -y @microsoft/powerbi-modeling-mcp@latest --start --readonly
```

Then:

```
claude mcp list
```

**Check:** `powerbi-modeling` is listed as **Connected**. The first time can take 20–30 seconds while it downloads.

## Step 7. Connect to the model

1. Open the sample report (`MeridianSupplyChain.pbix`) in Power BI Desktop and wait until the visuals have loaded. **Keep only one Power BI Desktop window open.**
2. In VS Code, **File › Open Folder** and pick the training folder (`Meridian_SupplyChain`).
3. Start a **new** Claude session: type `claude` in the terminal, or click **New session** in the Claude Code panel. Sessions that were already open don't see the new server.
4. Type `/mcp` and check **powerbi-modeling** is listed. Press **Esc**.

**5.** Type:

```
Connect to 'MeridianSupplyChain' in Power BI Desktop
```

Claude asks permission before each step. **Read the box**, then choose **Yes**.

**6.** Try three questions, and check each answer in Power BI Desktop:

- `How many tables are in the model, and what are they?`
- `List the measures with their DAX.`
- `List the relationships, and tell me which ones are inactive.`

**Check:** Claude says it's connected to MeridianSupplyChain, and its answers match Desktop (10 tables, 30 measures, 15 relationships with 3 inactive).

**Data safety:** read-only protects the **model**, not the **data**. Connecting doesn't upload anything, but every result Claude asks for (including rows from a DAX query) goes into the conversation and to the AI provider. Use training data only, and read every approval before you click.

## Step 8. Troubleshooting

| You see | Fix |
| --- | --- |
| `error: unknown option '-y'` | PowerShell swallowed the dashes. Put quotes around them: `'--'` (Step 6) |
| `'claude' is not recognized` | Close the terminal and open a new one (Step 4) |
| `powerbi-modeling` missing from `claude mcp list` or `/mcp` | Run the Step 6 command again. Already exists? `claude mcp remove powerbi-modeling`, then add it again |
| Server is listed, but Claude doesn't use it | Start a **new** Claude session. Open sessions keep their old server list |
| Sign-in says no access | Free plan. Claude Code needs Pro, Max, Team or Enterprise |
| Sign-in loops or the browser doesn't open | Set a default browser; press `c` in the terminal to copy the link and open it yourself; behind a proxy, ask IT (Part A, A3) |
| Server asks you to accept licence terms | Accept them once; it won't ask again |
| "No Power BI Desktop instance found" | Open the .pbix and wait until the visuals load, then ask again |
| Claude connects to the wrong report | Close all but one Power BI Desktop window |
| Claude says it can't create or change something | That's on purpose: the server is **read-only** (`--readonly`). This training starts read-only |

## Step 9. Final checklist before Session 2

- [ ] `node --version` shows v22 or higher
- [ ] `git --version` works, and your Git name and email are set
- [ ] Power BI Desktop can **Save as › Power BI project files (\*.pbip)**
- [ ] `claude --version` and `claude doctor` work
- [ ] You're signed in to Claude (terminal or VS Code)
- [ ] `claude mcp list` shows `powerbi-modeling` as Connected
- [ ] Claude connected to MeridianSupplyChain and answered the three test questions correctly

**Send your trainer:** a screenshot of Claude connected to MeridianSupplyChain, plus the output of:

```
claude doctor
claude mcp list
```

Stuck on a step? Send the screenshot of the error and the step number.
