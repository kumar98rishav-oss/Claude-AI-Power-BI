@echo off
rem Runs the report-snapshot Skill with nobody at the keyboard and appends the result to logs\snapshot.log.
rem Put this file in the project folder (next to the .pbip). Power BI Desktop must be open with the report.
cd /d "%~dp0"
if not exist logs mkdir logs
echo ===== %DATE% %TIME% ===== >> logs\snapshot.log
claude -p "/report-snapshot" --allowedTools "mcp__powerbi-modeling__*" "Bash(python *)" --max-turns 30 >> logs\snapshot.log 2>&1
