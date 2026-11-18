"""List the Power Query (M) steps behind every table, in plain English.

Usage:  python pq_steps.py <path to *.SemanticModel/definition>

Reads only the TMDL files. For each table: where the data comes from,
then each applied step with what it does.
"""
import re
import sys
from pathlib import Path

# M function -> plain-English description. Order matters: first match wins.
STEPS = [
    ("Sql.Database", "Connect to SQL Server"),
    ("Value.NativeQuery", "Run a native SQL query"),
    ("Excel.Workbook", "Read an Excel workbook"),
    ("Csv.Document", "Read a CSV file"),
    ("Json.Document", "Read a JSON file"),
    ("SharePoint.", "Read from SharePoint"),
    ("Folder.Files", "List files in a folder"),
    ("#table", "Inline table typed into the query (no source)"),
    ("Table.PromoteHeaders", "Use first row as headers"),
    ("Table.TransformColumnTypes", "Change data types"),
    ("Table.ReplaceValue", "Replace values"),
    ("Table.SelectRows", "Filter rows"),
    ("Table.SelectColumns", "Keep only some columns"),
    ("Table.RemoveColumns", "Remove columns"),
    ("Table.RenameColumns", "Rename columns"),
    ("Table.AddColumn", "Add a custom column"),
    ("Table.DuplicateColumn", "Duplicate a column"),
    ("Table.SplitColumn", "Split a column"),
    ("Table.TransformColumns", "Transform column values"),
    ("Table.NestedJoin", "Merge with another query"),
    ("Table.ExpandTableColumn", "Expand merged columns"),
    ("Table.Combine", "Append queries"),
    ("Table.Group", "Group by"),
    ("Table.Distinct", "Remove duplicates"),
    ("Table.Sort", "Sort rows"),
    ("Table.Unpivot", "Unpivot columns"),
    ("Table.Pivot", "Pivot a column"),
    ("Table.FillDown", "Fill down"),
    ("Table.Skip", "Skip top rows"),
    ("Table.FirstN", "Keep top rows"),
    ("Table.Buffer", "Buffer the table in memory"),
]


def describe(expr):
    for fn, text in STEPS:
        if fn in expr:
            if fn == "Table.TransformColumnTypes":
                n = len(re.findall(r'\{\s*"[^"]+"\s*,\s*[^{}]+?\}', expr))
                return f"{text} ({n} columns)"
            if fn == "Table.ReplaceValue":
                cols = re.findall(r'\{([^{}]*)\}\s*\)\s*,?\s*$', expr.strip())
                return f"{text} (blank to null in {cols[0].strip()})" if '""' in expr and "null" in expr and cols else text
            if fn in ("Csv.Document", "Excel.Workbook", "Json.Document"):
                f = re.search(r'"([^"]+\.(?:csv|xlsx|xlsm|xls|json))"', expr, re.I)
                return f"{text}: {f.group(1)}" if f else text
            return text
    head = re.match(r"\s*([A-Za-z#][\w.]*)", expr)
    return f"Other step ({head.group(1) if head else 'expression'})"


def parse_let(source):
    """Split 'let A = ..., B = ... in X' into [(name, expression)]."""
    body = re.search(r"\blet\b(.*)\bin\b", source, re.S)
    if not body:
        return [("Source", source.strip())]
    text, steps, depth, buf = body.group(1), [], 0, ""
    for ch in text:
        if ch in "([{":
            depth += 1
        elif ch in ")]}":
            depth -= 1
        if ch == "," and depth == 0:
            steps.append(buf)
            buf = ""
        else:
            buf += ch
    steps.append(buf)
    out = []
    for s in steps:
        if "=" in s:
            name, expr = s.split("=", 1)
            out.append((name.strip().strip('#"'), " ".join(expr.split())))
    return out


def partitions(tmdl_text):
    """Yield (partition name, mode, M source) from a table's TMDL."""
    lines = tmdl_text.splitlines()
    for i, line in enumerate(lines):
        m = re.match(r"\s*partition\s+(.+?)\s*=\s*m\s*$", line)
        if not m:
            continue
        mode, src, j = "import", [], i + 1
        while j < len(lines) and not re.match(r"\s*(partition|annotation|column|measure)\b", lines[j]) or (j < len(lines) and lines[j].startswith("\t\t\t")):
            l = lines[j]
            mm = re.match(r"\s*mode:\s*(\w+)", l)
            if mm:
                mode = mm.group(1)
            elif re.match(r"\s*source\s*=", l):
                src.append(l.split("=", 1)[1])
            elif src:
                src.append(l)
            j += 1
        yield m.group(1).strip("'"), mode, "\n".join(src).strip()


def main(definition):
    d = Path(definition)
    params = {}
    expr_file = d / "expressions.tmdl"
    if expr_file.exists():
        for m in re.finditer(r'expression\s+(\w+)\s*=\s*"([^"]*)"', expr_file.read_text(encoding="utf-8-sig")):
            params[m.group(1)] = m.group(2)
    if params:
        print("Parameters:", "; ".join(f"{k} = {v}" for k, v in params.items()), "\n")
    for f in sorted((d / "tables").glob("*.tmdl")):
        for name, mode, src in partitions(f.read_text(encoding="utf-8-sig")):
            steps = parse_let(src)
            print(f"## {f.stem}  (partition {name}, {mode} mode, {len(steps)} steps)")
            for i, (step, expr) in enumerate(steps, 1):
                print(f"  {i}. {step}: {describe(expr)}")
            print()


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
