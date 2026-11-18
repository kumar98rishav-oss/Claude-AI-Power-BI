"""Summarise a Power BI Performance Analyzer export, page by page.

Usage:  python perf_summary.py <PowerBIPerformanceData.json> <path to *.Report/definition>

Reads only. Prints, for each page, the page render time and every visual
sorted slowest first, split into DAX query / visual display / other.
Visual IDs in the export are the PBIR visual folder names, so the report
folder is used to name each page and visual.
"""
import json
import sys
from datetime import datetime
from pathlib import Path


def ts(s):
    return datetime.fromisoformat(s.replace("Z", "+00:00"))


def ms(e):
    return (ts(e["end"]) - ts(e["start"])).total_seconds() * 1000 if e.get("end") else 0.0


def load_pages(report_def):
    pages = {}
    pages_dir = Path(report_def) / "pages"
    order = json.loads((pages_dir / "pages.json").read_text(encoding="utf-8-sig")).get("pageOrder", [])
    for page_dir in sorted(p for p in pages_dir.iterdir() if p.is_dir()):
        meta = json.loads((page_dir / "page.json").read_text(encoding="utf-8-sig"))
        visuals = {}
        for vj in (page_dir / "visuals").glob("*/visual.json"):
            v = json.loads(vj.read_text(encoding="utf-8-sig"))
            vis = v.get("visual", {})
            title = None
            try:
                title = vis["visualContainerObjects"]["title"][0]["properties"]["text"]["expr"]["Literal"]["Value"].strip("'")
            except (KeyError, IndexError, TypeError):
                pass
            visuals[v["name"]] = {"type": vis.get("visualType", "?"), "title": title}
        pages[page_dir.name] = {"display": meta.get("displayName", page_dir.name), "visuals": visuals}
    rank = {name: i for i, name in enumerate(order)}
    return dict(sorted(pages.items(), key=lambda kv: rank.get(kv[0], 99)))


def main(perf_path, report_def):
    data = json.loads(Path(perf_path).read_text(encoding="utf-8-sig"))
    events = sorted(data["events"], key=lambda e: e["start"])
    by_id = {e["id"]: e for e in events}
    pages = load_pages(report_def)

    # Split the recording into segments: each user action (refresh, page change) starts one.
    segments, current = [], None
    for e in events:
        if e["name"] == "User Action":
            current = {"action": e["metrics"].get("sourceLabel", ""), "start": e["start"], "visuals": []}
            segments.append(current)
        elif e["name"] == "Visual Container Lifecycle" and current is not None:
            current["visuals"].append(e)

    def lifecycle_of(e):
        seen = 0
        while e and e["name"] != "Visual Container Lifecycle" and seen < 50:
            e = by_id.get(e.get("parentId"))
            seen += 1
        return e["id"] if e and e["name"] == "Visual Container Lifecycle" else None

    dax_ms, render_ms = {}, {}
    for e in events:
        if e["name"] in ("Execute DAX Query", "Render"):
            lc = lifecycle_of(e)
            if lc:
                bucket = dax_ms if e["name"] == "Execute DAX Query" else render_ms
                bucket[lc] = bucket.get(lc, 0.0) + ms(e)

    # Name each segment's page: the page whose visual folders contain every visual ID seen.
    latest = {}
    for seg in segments:
        ids = {v["metrics"]["visualId"] for v in seg["visuals"]}
        if not ids:
            continue
        match = [p for p, info in pages.items() if ids <= set(info["visuals"])]
        page = match[0] if len(match) == 1 else None
        seg["page"] = page
        if page:
            latest[page] = seg  # keep the most recent run of each page

    print("Recording:", Path(perf_path).name)
    print("Pages timed:", len(latest), "of", len(pages), "(latest run of each page)\n")
    rows_all = []
    for page, info in pages.items():
        seg = latest.get(page)
        if not seg:
            print(f"## {info['display']}: not in this recording (open the page while recording)\n")
            continue
        vs = seg["visuals"]
        page_ms = (max(ts(v["end"]) for v in vs) - min(ts(v["start"]) for v in vs)).total_seconds() * 1000
        print(f"## {info['display']}: {page_ms:,.0f} ms to render ({len(vs)} visuals, loaded in parallel)")
        print(f"{'#':>2}  {'Visual':<34} {'Type':<18} {'Total ms':>9} {'DAX ms':>7} {'Display ms':>10} {'Other ms':>8}")
        rows = []
        for v in vs:
            vid = v["metrics"]["visualId"]
            meta = info["visuals"].get(vid, {})
            name = meta.get("title") or v["metrics"].get("visualTitle") or vid
            total, dax, disp = ms(v), dax_ms.get(v["id"], 0.0), render_ms.get(v["id"], 0.0)
            rows.append((total, f"{name} [{vid}]", v["metrics"].get("visualType", meta.get("type", "?")), dax, disp))
        rows.sort(reverse=True)
        for i, (total, name, vtype, dax, disp) in enumerate(rows, 1):
            other = max(total - dax - disp, 0)
            print(f"{i:>2}  {name[:34]:<34} {vtype[:18]:<18} {total:>9,.0f} {dax:>7,.0f} {disp:>10,.0f} {other:>8,.0f}")
            rows_all.append((total, info["display"], name))
        print()
    if rows_all:
        rows_all.sort(reverse=True)
        print("Slowest visuals across the report:")
        for total, page, name in rows_all[:3]:
            print(f"  {total:,.0f} ms  {page} / {name}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2])
