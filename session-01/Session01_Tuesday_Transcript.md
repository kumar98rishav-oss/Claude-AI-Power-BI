# Session 01 Tuesday Transcript

Your minute-by-minute script for **Connecting Claude to Power BI**: 60 minutes, 20 slides. The quote blocks are what you say, so read them like you're chatting, not presenting. **Do** lines are what you click.

Live, editable version: https://claude.ai/artifact/KahHMqz8xB9U8MsqiuT3rM

## Before Tuesday

### The day before

**1.** On your own laptop, work from the **`1st Session`** folder (in this repo: `pbip-project/`). It's the PBIP copy of Meridian, and the report-snapshot Skill lives there. Open it in VS Code (**File → Open Folder**), open a terminal with **Ctrl + \`**, and check it says **PowerShell**.

**2.** Check your installs:

```
node --version
claude --version
claude doctor
```

**3.** Register the Power BI server once for your user account, so it works in any folder:

```
claude mcp add --scope user powerbi-modeling '--' cmd /c npx -y @microsoft/powerbi-modeling-mcp@latest --start --readonly
claude mcp list
```

If it says the name already exists, run `claude mcp remove powerbi-modeling` and add it again. If the server asks you to accept licence terms, do that now so it doesn't pop up mid-class.

**4.** Make sure `PowerBIPerformanceData.json` is sitting next to `MeridianSupplyChain.pbip`. If you've changed the report since it was recorded, record it again: **Optimize › Performance analyzer › Start recording › Refresh visuals**, click through the other two pages, then **Stop › Export**.

**5.** Do one full run-through with a stopwatch: connect, the three hands-on questions, the Pacific Rim breakdown, and `/report-snapshot`.

### 30 minutes before

- Open these, in this order: **PowerPoint** (Slide Show, slide 1); **Power BI Desktop** with `MeridianSupplyChain.pbip` on Executive Overview, Year = 2025; **VS Code** on the `1st Session` folder with a terminal open. It's the same report as the .pbix everyone else has.
- Keep only **one** Power BI Desktop open, so Claude doesn't have to guess which one you mean.
- Practise Alt+Tab between PowerPoint, Desktop and VS Code a couple of times.
- Zoom VS Code and the terminal (**Ctrl + =**) until the back row can read them. Notifications off.

### Message to send everyone the day before

> Hey everyone! Quick prep for Tuesday, about 10 minutes:
>
> 1. Install **Power BI Desktop**, **VS Code**, and **Node.js 22 or later** (nodejs.org).
> 2. Copy the **Meridian_SupplyChain** folder to `C:\Training\` and open `MeridianSupplyChain.pbix` once in Power BI Desktop.
> 3. In VS Code, open a terminal (Ctrl + \`) and run: `irm https://claude.ai/install.ps1 | iex`
> 4. Close that terminal, open a new one, and type `claude` to sign in. You'll need a Claude **Pro, Max, Team or Enterprise** account; the free plan doesn't include Claude Code.
> 5. Check these three work: `node --version`, `claude --version`, `claude doctor`.
>
> If anything breaks, just send me a screenshot before Tuesday. See you then!

Every person who shows up already set up saves the room about ten minutes.

## 0:00–15:00 · Slides 1–6 · Hello, and the big ideas

### 0:00–2:00 · Slide 1: Connecting Claude to Power BI

**Screen:** slide 1. **Do:** give people a minute to join, then start.

> Hi everyone, good evening! Thanks for joining.
>
> Quick intro first: I'm Rishabh, and I work as a Data Engineer.
>
> Over the next few sessions, we're going to look at how Claude can plug into Power BI and take some of the everyday grind off our plates.
>
> Because let's be honest: building the report is only part of the job. Keeping it running, making sure the data's right, looking after the model, chasing down why something broke... that's where a lot of our time actually goes.
>
> So in these sessions I'll share practical, real-world ways to make our Power BI work faster and easier with a bit of help from AI. Nothing theoretical; stuff you can use the next morning.
>
> *(Point at the four boxes.)* Here's today in four steps: what MCP is, install one thing, connect Claude to a report read-only, then ask it questions and check the answers.
>
> And if you remember one word from today, make it **check**. We never trust an answer we haven't checked ourselves.
>
> Alright, let's get started with session one!

### 2:00–3:00 · Slide 2: Where these sessions take us

> Before we jump in, here's where we're heading over the next four or five sessions. *(Point at each card.)*
>
> **One: automate the routine reporting work.** Answering questions about a model, running checks, repeat changes. Claude does the legwork, you approve and check.
>
> **Two: really understand Power BI's metadata.** How a report and its model are actually stored, as PBIP, TMDL and PBIR, and how that lets us rebuild things instead of redoing them by hand.
>
> **Three: build Skills on top of TMDL.** Once we can read those files, we write Skills that work on them, so a procedure runs the same way every time, whoever runs it.
>
> **Four: troubleshoot any report.** Skills that explain a report, find what's broken and why, so checks that take hours take minutes.
>
> *(Point at the amber strip.)* Today is step one: connect, check, and build our first Skill before the hour's up.

**Ask:** "Quick one: has everyone got Power BI Desktop open? Thumbs up." Anyone without it: "No worries, pair up with your neighbour for today."

### 3:00–5:00 · Slide 3: Meet Meridian Medical Supply

> We need a real model to play with, so meet Meridian Medical Supply. Made-up company, real-looking data. They distribute medical supplies, gloves, masks, test kits, IV sets, from five warehouses to about a hundred and fifty hospitals and clinics. Two years of data.

**Do:** Alt+Tab to Power BI Desktop. Click the Year slicer: 2024, then 2025.

> Here's the report. Revenue, margin, growth, on-time delivery. Flip the year and everything moves.

**Do:** Alt+Tab back to the slide.

> Behind it: three fact tables, thirty measures, fifteen relationships. If you haven't yet, open MeridianSupplyChain.pbix now and just leave it open. Claude's going to connect to exactly what you're looking at.
>
> And fair warning: there are a few real business problems hiding in this data. We'll dig one out together later.

### 5:00–9:00 · Slide 4: How Claude reaches a Power BI model

> So how does Claude actually get inside a Power BI model? Follow me left to right.
>
> **You** type a question in plain English, into **Claude Code** inside VS Code. That's the host. Inside it is an **MCP client**, basically a connection that knows how to talk to one server. The **Power BI MCP server** is a small program from Microsoft running on your own laptop. It finds the Power BI Desktop you've got open and talks to the **model** inside it, the same engine DAX Studio and Tabular Editor connect to.

**Ask:** "Before today, what would you have used to poke around inside a running model?" Give it five seconds. You'll hear DAX Studio or Tabular Editor.

> Exactly. MCP gives Claude that same kind of access.
>
> *(Point at the four steps.)* Here's what happens when you ask something. One: you ask, "Which supplier was late most often in 2025?" Two: Claude picks a tool, say, run a DAX query. Three, and this is the big one: **you approve**. Claude shows you the tool and what it's about to send before anything runs. Four: the model answers. Pacific Rim Medical, sixty-one percent on time. So Claude isn't guessing. The number comes straight from the model.

### 9:00–12:00 · Slide 5: One standard plug between AI and tools

> MCP stands for Model Context Protocol. It's an open standard, put out by Anthropic in November 2024, for hooking AI apps up to tools and data.
>
> *(Point at the amber box.)* Think of it like a universal plug. Without it, every AI app needs custom code for every tool. With MCP, you build the tool once, as a server, and any app that speaks MCP can use it. That's why Microsoft only had to build one Power BI server, and it works with Claude Code, Claude Desktop and others.
>
> *(Point at the three cards.)* A server can offer three things. **Tools**: actions Claude can ask to run, like list tables or run a DAX query. **Resources**: data Claude can read. **Prompts**: ready-made instructions. Today it's all tools.
>
> When Claude connects, the server hands over its tool list: names, what each does, what it needs. Claude can only use what's on that list. I'll show you the actual list in a few minutes.

### 12:00–15:00 · Slide 6: Inside the server: TOM and ADOMD.NET

> *(Point at the left box.)* So what's the server actually doing under the hood? It's a .NET program built on two Microsoft libraries. Both are bundled inside it; I checked.
>
> *(Point at the purple box.)* First one's **TOM**, the Tabular Object Model. It reads and changes the model's structure: tables, columns, measures and their DAX, relationships. If you've used Tabular Editor, you've used TOM. Claude listing tables or showing you a measure's DAX? That's TOM reading. Claude creating a measure? That's TOM writing, and that's exactly what read-only mode blocks.
>
> *(Point at the amber box.)* Second one's **ADOMD.NET**. It runs DAX queries and brings the rows back. Used DAX Studio? Then you've used ADOMD.NET. Every number Claude gives you, 2025 revenue, a supplier's on-time rate by quarter, comes back through ADOMD.NET.
>
> *(Point at the right box, then the table.)* Both plug into the same thing: the Analysis Services engine inside Power BI Desktop, on a local port. So MCP isn't some new back door into your model. It's the same access Tabular Editor and DAX Studio already have, and it asks you first, every time.

**Ask:** "When Claude says 2025 revenue was 14.1 million, which library fetched that?"

**Answer:** "ADOMD.NET, because it's a DAX query. Listing the tables would be TOM."

**If someone asks how we know:** the DLLs Microsoft.AnalysisServices.Tabular (TOM) and Microsoft.AnalysisServices.AdomdClient (ADOMD.NET) are inside the server's exe. Which tool uses which library is my inference from what each one is built for; Microsoft doesn't document that.

> Okay, that's the theory out of the way. Let's get it set up.

## 15:00–30:00 · Slides 7–11 · Setup and connect

### 15:00–18:00 · Slide 7: Open the project, then check what's installed

**Screen:** slide 7, then VS Code.

> Before we install anything, let's open the right folder and see what you've already got.
>
> *(Point at the left cards.)* In VS Code: File, Open Folder, pick Meridian_SupplyChain. Then open a terminal: Control plus the backtick key, the one under Escape. Check the terminal panel says PowerShell. And keep MeridianSupplyChain.pbix open in Power BI Desktop.

**Do:** show it on your screen (you're in `1st Session`, which is the same report). Then:

> Now type these three with me, one at a time.

```
node --version
claude --version
claude doctor
```

> *(Point at the table.)* Here's how to read your screen. Node 22 or higher, you're good. If claude --version shows something like 2.1 and "Claude Code", Claude's installed. And claude doctor with no errors means it's healthy.
>
> If any of them says "is not recognized", that bit just isn't installed yet. Hands up, and we'll sort it on the next slide.

### 18:00–21:00 · Slide 8: Install Claude Code and sign in

**Only for the people whose check failed.** If everyone did the pre-work, run through this in a minute.

> If claude wasn't recognised, it's three quick steps. *(Point at each.)*
>
> One: paste this into the PowerShell terminal. It's Anthropic's official installer, and you don't need admin rights.

```
irm https://claude.ai/install.ps1 | iex
```

> Two: close the terminal with the bin icon and open a fresh one, Control and backtick again. The new terminal is the one that knows about the claude command. Then check:

```
claude --version
```

> Three: type claude and hit Enter. A browser window pops up. Sign in with your Claude account, then come back to the terminal.

```
claude
```

> *(Point at the Account card.)* One heads-up: Claude Code needs a Pro, Max, Team or Enterprise account. The free plan doesn't include it.
>
> And if you've got the Claude extension in VS Code, keep it. It's just that the extension doesn't add the claude command to your terminal, and we need that command next.

**Do:** once they're signed in, show typing `/exit` to leave Claude so the terminal's free again.

**If someone's in CMD** (they see `'irm' is not recognized`): point at the "In CMD" line at the bottom of the slide, or switch their terminal to PowerShell.

### 21:00–25:00 · Slide 9: Register the Power BI server, read-only

> Everyone's got Claude. Now we tell Claude Code about the Power BI server, with one command. But let's read it before we run it, because reading before running is a habit worth having.
>
> *(Point at each card.)* claude mcp add: register a server. Dash-dash scope user: register it once for your account, so it works in every folder. powerbi-modeling: just the name we're giving it. The two dashes are in quotes on purpose. In PowerShell, two bare dashes can get swallowed before Claude sees them, and you'd get an error. Everything after that is how to start the server: cmd slash c npx downloads and runs Microsoft's package.
>
> And the bit at the end: dash-dash readonly. Microsoft calls it safe mode. Claude can't change anything in your model. Today, Claude can look, but not touch.

**Do:** Alt+Tab to VS Code. "Type this with me":

```
claude mcp add --scope user powerbi-modeling '--' cmd /c npx -y @microsoft/powerbi-modeling-mcp@latest --start --readonly
```

> Hit Enter. Now let's check it worked.

```
claude mcp list
```

**Expected:** `powerbi-modeling` listed as connected. The first run takes 20–30 seconds while it downloads.

> Thumbs up when you see powerbi-modeling in the list.

### 25:00 · Slide 10: Quick fixes (only if someone's stuck)

> If something went wrong, find your error on the left. The fix is right next to it.

The two you'll see most: `'claude' is not recognized` (they didn't open a new terminal after installing) and `unknown option '-y'` (the dashes weren't in quotes). Pair anyone still stuck with a neighbour and keep moving.

### 25:00–30:00 · Slide 11: Connect Claude to the open report (LIVE)

> Okay, this is the fun bit. We're connecting Claude to the report that's open in Power BI Desktop.

**Do:** Alt+Tab to Desktop and show the report's loaded.

> Quick check first: MeridianSupplyChain open in Desktop, and fully loaded, not sitting on the start page.

**Do:** Alt+Tab to VS Code. In the terminal:

```
claude
```

> That starts Claude Code. Now type this sentence exactly. It's straight out of Microsoft's own docs.

```
Connect to 'MeridianSupplyChain' in Power BI Desktop
```

**Do:** when the first approval box pops up, **stop. Don't click yet.**

> Look at this. Claude's asking permission. It wants the Power BI server's connection tool, to list local instances. In plain English: can I look for Power BI Desktop on this laptop? Totally harmless, so we say yes.

**Do:** choose **Yes**, not "don't ask again".

> Second request: now it wants to connect, using the address it just found. Yes again.
>
> And there we go. It found one Power BI Desktop called MeridianSupplyChain and connected. Your port number will be different from mine; that's normal.
>
> What we just did, reading the box before clicking, honestly that's the most important habit in this whole programme.

**Do (30 seconds):** type `/mcp`, open **powerbi-modeling**.

> Remember on slide 5 I said the server hands Claude a list of tools? Here it is: connection, tables, columns, measures, relationships, DAX queries. Claude can only use what's on this list. The table, column, measure and relationship tools go through TOM; the DAX query tool goes through ADOMD.NET. The two libraries from slide 6.

**Do:** press Esc.

> Everyone connected? Thumbs up.

**If someone gets "no instance found":** their .pbix hasn't finished loading. Wait for the visuals, then ask again.

## 30:00–45:00 · Slides 12–13 · Reading the model, and your turn

### 30:00–33:00 · Slide 12: Four things Claude can read

> So Claude's connected. What can it actually see? The four building blocks of every model. *(Point at each card.)* Tables: ten in Meridian. Columns: eighty-four, with their data types. Measures: thirty, with the full DAX. Relationships: fifteen, including which side is "many" and which ones are switched off.
>
> Under each card there's a question you can ask in plain English, and the tool Claude will reach for. You don't need to memorise tool names, but you'll see them in the approval box, so it helps to recognise them.

**Do:** Alt+Tab to VS Code and type:

```
Show me the DAX for Gross Margin % and the table it lives in.
```

Approve the tool call, reading it out: "Measure operations, get. That's just a read."

**Expected:** table `_Measures`, folder *01 Sales*, DAX `DIVIDE ( [Gross Margin], [Revenue] )`, format `0.0%`.

> That's Claude's answer. But we don't take its word for it until we check.

**Do:** Alt+Tab to Desktop. In the **Data** pane, expand `_Measures`, click **Gross Margin %**, show the formula bar.

> Same table, same formula. That's all checking is. Now you try it, three times.

### 33:00–34:30 · Slide 13: Explain the hands-on

**Screen:** slide 13. Leave it up for the whole exercise.

> Your turn! Ask Claude these three questions, one at a time.
>
> *(Point at the box on the right.)* And here's the rule that matters more than the questions: write Claude's answer down first, on paper or in Notepad. Then find the same thing in Power BI Desktop. Then compare.
>
> Why write it down first? Because once you've seen the right answer in Desktop, it's really easy to convince yourself Claude got it right too. Writing it first keeps you honest.
>
> Each question tells you where to check. You've got eight minutes, and I'll be walking around. Go for it!

**Do:** start an 8-minute timer where everyone can see it.

### 34:30–42:30 · Work time (you walk around)

Keep these answers to yourself until the debrief:

| # | Question | Right answer | Where to check in Desktop |
| --- | --- | --- | --- |
| 1 | How many tables, measures and relationships? | 10, 30, 15 | Data pane; Modeling → Manage relationships |
| 2 | Which relationships are inactive, and why? | Sales_ShipDate, Purchases_PromisedDate, Purchases_ReceivedDate. Several date columns point at one date table, and only one link can be active | Model view: dotted lines |
| 3 | 2025 revenue, and the worst supplier on-time %? | $14,134,114.90; Pacific Rim Medical, 61.2% | Executive Overview and Supplier Performance pages, Year = 2025 |

**Things to say as you walk around:**

- To someone stuck: "Read me the approval box. What's it asking to do?"
- To someone quick: "Did it match? Nice. Now ask Claude how it worked it out."
- To someone whose numbers don't match: "Ooh, good, that's the interesting one. Hang on to it for the debrief."

**At 7 minutes:** "One more minute, just finish the one you're on."

### 42:30–44:00 · Debrief

> Okay, pens down. Let's compare. Question one?

Confirm: "Ten, thirty, fifteen."

> Question two: the inactive relationships?

**Do:** Alt+Tab to Desktop, click **Model view**, and point at the dotted lines.

> There they are, the dotted lines. A sales order has an order date and a ship date, but the date table can only have one active link to each fact table. The others stay switched off until a measure turns them on with USERELATIONSHIP.
>
> Question three: 2025 revenue, and the worst supplier?

Confirm: "Fourteen point one million. Pacific Rim Medical, sixty-one percent."

**Ask:** "Did anyone get something different from Claude? What happened?" Take one or two.

> And that's exactly why we check. On metadata Claude's usually spot on, and a DAX query result is exact. But it can misread the question, like counting a hidden table, or using order date when you meant ship date. Claude's fast. You're the check.

### 44:00–45:00 · The story hook

> One more thing. Pacific Rim Medical is at sixty-one percent for the year. Was it always that bad? Let's ask.

**Do:** on your screen:

```
Break down Pacific Rim Medical's on-time % by quarter for 2024 and 2025.
```

Approve the DAX query, reading it out.

**Expected:**

| Quarter | 2024 | 2025 |
| --- | --- | --- |
| Q1 | 100% | 88.9% |
| Q2 | 87.0% | 88.9% |
| Q3 | 96.0% | **28.0%** |
| Q4 | 96.2% | **33.3%** |

> Look at that! Around ninety percent on time for six quarters straight. Then from July 2025 it just falls off a cliff: twenty-eight percent, then thirty-three.
>
> That's not just a number on a dashboard any more. That's a real business problem, a supplier whose deliveries collapsed. And we found it with one sentence, and we can check it in Desktop. That's the real payoff of what you set up today.
>
> Now, you've typed a few questions. Imagine running the same checks on every report, every morning. Let's save them once, as a Skill.

If anyone asks: the quarters follow the purchase order date, which is the model's active relationship.

## 45:00–60:00 · Slides 14–20 · Skills, files, rules and wrap-up

### 45:00–47:00 · Slide 14: What a Skill is

> So, you've just typed three questions and checked three answers. Now picture doing that on every report, every morning, or before every release. You'd write it down as a checklist so anyone could follow it, right? A Skill is exactly that, just written for Claude.
>
> *(Point at the green band.)* A Skill is a saved procedure. You write the steps once, in plain English, and Claude follows them the same way every time, whoever runs it.
>
> *(Point at the code box.)* It's literally a folder with one file in it, SKILL.md. Three parts: a name, which is what you call it; a description, which tells Claude when to use it; and the steps, written the way you'd explain them to a new teammate.
>
> *(Point at the two cards.)* Where you save it decides who gets it. Inside the project, in .claude/skills, it travels with the project through Git, so the whole team gets it. In your home folder, it's just yours, in every project.
>
> *(Point at the amber box.)* And to use it, you call it by name with a slash, or just ask in your own words.

### 47:00–49:00 · Slide 15: What Skills can do for a Power BI team

**Keep it light.** A sentence or two per card; each one gets its own session later.

> Here's why Skills are a big deal for anyone building or checking Power BI reports every day. Six jobs most of us do by hand right now.
>
> **1. Daily health check.** Every morning: are the key measures returning values, with no errors? Anything missing a description or a format?
>
> **2. Source against report.** Right now Meridian runs off files. Soon we'll move the same data into SQL Server and hook up a SQL Server MCP next to the Power BI one. Then a Skill can ask the same question twice, say revenue by month from SQL and from the model, and flag any month that doesn't match.
>
> **3. Threshold alerts.** You put the thresholds in the Skill: revenue above target, a supplier's on-time rate below 80%. Open the report, run the Skill, and it tells you what crossed a line. Just to be clear, it runs when you run it. It's not a push notification.
>
> **4. What changed.** Once the report's saved as PBIP, which is next session, every change is a file change. A Skill reads the Git diff and tells you which measures, columns, relationships or visuals changed.
>
> **5. Visual and formatting check.** The report pages are files too. A Skill can spot visuals pointing at fields that don't exist any more, missing titles, and colours or fonts that don't match the theme.
>
> **6. Performance.** A Skill can list the biggest columns and time the slowest measures. And once we're on SQL Server, it can read the Power Query steps and flag the ones likely to break query folding. You confirm with View Native Query: Claude points to the step, you check it.
>
> *(Point at the amber band.)* Today's just the idea. We'll build these one by one.

**Ask:** "Which of these six would save you the most time this week?" Take one or two answers, then move on.

**If someone asks:** the SQL Server MCP and PBIP aren't part of today's setup. Cards 1 and 3, and the model side of card 6, work with today's read-only connection.

### 49:00–51:00 · Slide 16: A Skill for every morning: report-snapshot

> Okay, here's a real one, built for this report, and I'm going to run it in a minute. It's called report-snapshot, and it answers five questions you'd otherwise check by hand every morning.
>
> *(Point at each row.)* One: when was the data last refreshed? Two: what are the key KPIs this year against last year, and has anything crossed a threshold we set? Three: how many relationships are there, what kind, and which are switched off? Four: how long does each page take to load, and which visual is the slowest? Five: what did Power Query do to each table?
>
> *(Point at the tags on the right.)* Each answer comes from a different place. One to three are DAX queries through the Power BI server, so that's ADOMD.NET from slide 6. Four comes from Performance Analyzer: you record it once in Desktop and export it, because Claude can't click around in Desktop for you. Five comes from the model's TMDL files.
>
> *(Point at the dark panel.)* And this is what it actually found on Meridian. Two flags: on-time delivery is 89.8%, just under our 90% line, and there are 45 stockouts, over our 40 line.

**If someone asks:** the Skill is a SKILL.md plus two small Python scripts that read the Performance Analyzer file and the Power Query steps, so those parts come out the same every run. The thresholds are a little table in SKILL.md you can edit for your team.

### 51:00–55:00 · Slide 17: Save, recall and run a Skill (LIVE)

**Already set up:** VS Code on `1st Session`, `MeridianSupplyChain.pbip` open in Desktop, and `PowerBIPerformanceData.json` next to the .pbip.

**Do:** in VS Code, open `.claude/skills/report-snapshot/SKILL.md` and scroll through it: the description, the read-only rule, the five steps, the threshold table, the output format.

> This is honestly all a Skill is: one file of instructions, plus a couple of small scripts. It's saved in the project, so it goes wherever the project goes and the whole team gets it. Want your own? Ask Claude Code to draft one, read what it wrote, then save it.

**Do:** restart Claude Code so it picks up the Skill, then call it:

```
/exit
claude
/report-snapshot
```

If `/report-snapshot` doesn't show up, type `Give me today's report snapshot` instead. Approve each read-only call, and read the DAX in each approval out loud.

**Expected** (checked on the live model; render times wobble a little each recording):

| Check | Result |
| --- | --- |
| Last refreshed | 9 Nov 2026, 12:25 UTC (17:55 IST) |
| KPIs, 2025 vs 2024 | Revenue 14,134,115, up 10.1%. **Flags:** On-Time Delivery 89.8% (below 90%), Stockouts 45 (above 40) |
| Relationships | 15: 12 active, 3 inactive (Sales_ShipDate, Purchases_PromisedDate, Purchases_ReceivedDate), all many-to-one, single direction |
| Render times | Executive Overview 1,060 ms, Supplier Performance 1,012 ms, Inventory Health 792 ms. Slowest: KPI card, 954 ms |
| Power Query | 9 tables read CSV files, promote headers, change types. Fact_Sales and Fact_Purchases also turn blanks into nulls |

> Same five checks, same format, every morning, for anyone who opens this project. That's the whole point of a Skill.

### 55:00–56:00 · Slide 18: Underneath, the model as files

> Today Claude read the model live, straight from Desktop. Next session we go one level deeper.
>
> *(Left box.)* The .pbix you opened today is one sealed file: model, data and report all bundled together. *(Middle box.)* But Power BI can also save the same report as a project, a PBIP, and then it becomes a folder of plain text files: one folder for the report, in a format called PBIR, and one for the model, in a format called TMDL, with a text file per table. You actually just saw it: the Skill read those TMDL files.
>
> *(Bottom strip.)* Why should you care? Text files can be read by Claude, compared in Git, and undone with one command.
>
> *(Amber box.)* If you're curious before next week, these links open the interactive map: every file, what's in it, and what's safe to touch.

Keep this to one minute. The details are Session 2.

### 56:00–58:00 · Slide 19: Four rules

> Four rules, and they start today. One: read-only first. On any new model, start read-only. Write access is something you decide later, on purpose. Two: read every approval. The tool name and its inputs tell you exactly what's about to happen. Three: check every answer against Desktop, like you did today. Four: data leaves your machine. *(Point at the quote.)* That's Microsoft's own warning: what Claude reads can be sent to the AI provider. So stick to training data or approved models. Never a client's model without permission. And always follow our AI policy.

### 58:00–60:00 · Slide 20: Next, the model as files

> So, three things to take away. *(Point at each card.)* MCP is the plug: Claude reaches your model through a server running on your own laptop. Read-only first, and read every approval before you click. And check every answer against Desktop before you believe it.
>
> Next session, we save Meridian as a PBIP project, open the folder in VS Code, and ask Claude to rename a column across every measure that uses it. Then we'll read the change as a Git diff and undo it with one command.
>
> Before then, two optional things: ask Claude three more questions about Meridian and check them, and have a poke around the interactive map if you're curious.
>
> Thanks so much, everyone! Any questions?

**Do:** paste the map link in the chat and leave this slide up: https://rishav-portfolio-i46g.onrender.com/training/power-bi-internals-map.html

## Fixes, timing and numbers

### If something goes wrong

| You see | Fix |
| --- | --- |
| `'claude' is not recognized` | Install with the PowerShell line on slide 8, then open a **new** terminal |
| `'irm' is not recognized` | They're in CMD: use the CMD line on slide 8, or switch the terminal to PowerShell |
| `error: unknown option '-y'` | The dashes need quotes: `'--'` |
| `'node'` or `'npx'` is not recognized | Install Node.js 22 or later from nodejs.org, then restart VS Code |
| `powerbi-modeling` not connected | `claude mcp remove powerbi-modeling`, then run the add command again |
| Sign-in says no access | They're on the free plan. Claude Code needs Pro, Max, Team or Enterprise |
| Server blocks every tool call | Accept the licence terms it shows the first time |
| No Power BI Desktop instance found | Open the .pbix and wait until the visuals have loaded |
| `/report-snapshot` not recognised | Ask "Give me today's report snapshot". Check Claude was started in the `1st Session` folder |
| Snapshot says there's no performance recording | Record in Desktop (Optimize › Performance analyzer › Record › visit each page › Export) and save it next to the .pbip |

### Timing checkpoints

| Slide | Should start at | If you're running late |
| --- | --- | --- |
| 7 (setup) | 15:00 | Cut the Desktop demo on slide 3 to 20 seconds, and slide 6 (TOM and ADOMD.NET) to 30 seconds |
| 12 (four things) | 30:00 | Skip the `/mcp` tool list on slide 11 |
| 14 (Skills) | 45:00 | Cut hands-on work time to 6 minutes, but keep the story hook |
| 17 (live Skill) | 51:00 | Skip the restart: ask "Give me today's report snapshot" in words |
| 20 (wrap-up) | 58:00 | Slide 18 in thirty seconds |

**Never cut:** reading the approvals on slide 11, the hands-on, and the Pacific Rim story.

### Numbers to know by heart

- 10 tables, 30 measures, 84 columns, 15 relationships, 3 inactive
- 2025 revenue $14.13M, up 10.1% on 2024; Pacific Rim Medical 61.2% on time in 2025
- Pacific Rim by quarter: about 90% for six quarters, then **28%** and **33%** in the second half of 2025
- Gross Margin %: `_Measures`, `DIVIDE ( [Gross Margin], [Revenue] )`, `0.0%`
- report-snapshot flags: On-Time Delivery 89.8% (line is 90%), Stockouts 45 (line is 40); slowest visual the KPI card, 954 ms
