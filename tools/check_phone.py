"""Prove a page is phone-ready: render it in a real narrow frame, measure it, shoot it.

A small headless window is silently cropped, so the page is loaded inside an <iframe> of an exact
width on a much larger wrapper page. The wrapper reaches into the frame (same origin, thanks to
--allow-file-access-from-files) and measures what actually rendered, then prints the numbers.

Run (from the repo root):
    python tools/check_phone.py Alphabet/index.html              # 360px and 1200px
    python tools/check_phone.py Alphabet/index.html 360          # one width only

What it checks, per CLAUDE.md:
  * the page itself never scrolls sideways (scrollWidth <= clientWidth)
  * only a .tw table box may scroll sideways, and its first column must be sticky
  * every chart is a sensible size and no wider than 520px
  * every link and button is at least 40px tall
  * nothing sticks out past the right edge of the body

Needs Google Chrome. Writes a PNG per width next to the page, under tools/shots/.
"""
import json                                            # parses the measurements back out of the DOM
import re                                              # pulls the measurement block out of the dumped DOM
import subprocess                                      # runs headless Chrome
import sys                                             # reads the page path and widths from the command line
import tempfile                                        # a throw-away folder for the wrapper and the profile
from pathlib import Path                               # paths that work on Windows and Linux

ROOT = Path(__file__).resolve().parent.parent          # the repo root (this file is in tools/)
CHROME = Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe")   # Chrome on this machine
SHOTS = ROOT / "tools" / "shots"                       # where the screenshots are written

# The JavaScript the wrapper runs once the frame has loaded. It returns one plain object.
MEASURE = r"""
function measure(doc, win) {
  var el = doc.scrollingElement || doc.documentElement;
  var bodyW = doc.body.getBoundingClientRect().width;
  var charts = [];
  doc.querySelectorAll('svg[data-chart]').forEach(function (s) {
    var r = s.getBoundingClientRect();
    charts.push({ id: s.dataset.chart, w: Math.round(r.width), h: Math.round(r.height) });
  });
  var small = [];
  doc.querySelectorAll('a, button').forEach(function (a) {
    var r = a.getBoundingClientRect();
    if (r.height > 0 && r.height < 40) {
      small.push({ tag: a.tagName, h: Math.round(r.height * 10) / 10, text: (a.textContent || '').trim().slice(0, 40) });
    }
  });
  var scrollers = [];
  doc.querySelectorAll('.tw').forEach(function (t) {
    var tbl = t.querySelector('table');
    var first = tbl ? tbl.querySelector('tbody td') : null;
    scrollers.push({
      scrolls: t.scrollWidth > t.clientWidth,
      sticky: first ? win.getComputedStyle(first).position : 'none'
    });
  });
  var overflow = [];
  doc.querySelectorAll('body *').forEach(function (n) {
    var r = n.getBoundingClientRect();
    if (r.width > 0 && r.right > bodyW + 1 && !n.closest('.tw')) {
      overflow.push({ tag: n.tagName, cls: (n.className.baseVal || n.className || '') + '', right: Math.round(r.right) });
    }
  });
  return {
    scrollWidth: el.scrollWidth, clientWidth: el.clientWidth, bodyWidth: Math.round(bodyW),
    charts: charts, smallTargets: small, scrollers: scrollers, overflow: overflow.slice(0, 8)
  };
}
"""


def wrapper_html(page_url, width):
    """Build the wrapper page: a frame of an exact width, plus the measuring script."""
    return f"""<!doctype html><meta charset="utf-8"><title>frame</title>
<style>html,body{{margin:0;background:#444}} iframe{{border:0;display:block}}</style>
<iframe id="f" src="{page_url}" style="width:{width}px;height:3000px"></iframe>
<pre id="out" style="color:#fff;font:12px monospace"></pre>
<script>{MEASURE}
window.addEventListener('load', function () {{
  setTimeout(function () {{                       // give the fonts and the charts a moment to settle
    var f = document.getElementById('f');
    var r;
    try {{ r = measure(f.contentDocument, f.contentWindow); }}
    catch (e) {{ r = {{ error: String(e) }}; }}
    document.getElementById('out').textContent = 'RESULT' + JSON.stringify(r) + 'END';
  }}, 1200);
}});
</script>"""
    # returns the wrapper page as an HTML str, with the measuring script inlined


def run_chrome(args, work):
    """Run headless Chrome once with the given extra arguments."""
    base = [str(CHROME), "--headless=new", "--disable-gpu",
            f"--user-data-dir={work / 'profile'}",     # its own profile: the default one is locked
            "--allow-file-access-from-files",          # lets the wrapper read inside the frame
            "--hide-scrollbars", "--virtual-time-budget=9000"]
    return subprocess.run(base + args, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", timeout=180)
    # returns the finished subprocess.CompletedProcess


def check(page, width, work):
    """Render the page at one width, measure it, and report every rule it breaks."""
    url = (ROOT / page).resolve().as_uri()             # the page under test, as a file:// URL
    wrap = work / f"wrap{width}.html"                  # the wrapper that holds the frame
    wrap.write_text(wrapper_html(url, width), encoding="utf-8")

    dom = run_chrome(["--dump-dom", "--window-size=1600,3200", wrap.as_uri()], work)
    found = re.search(r"RESULT(\{.*?\})END", dom.stdout, re.S)   # the block the wrapper printed
    if not found:
        print(f"  {width}px: could not read measurements")
        return 1
    m = json.loads(found.group(1))                     # the measurements, as a plain dict
    if "error" in m:
        print(f"  {width}px: measuring failed: {m['error']}")
        return 1

    SHOTS.mkdir(parents=True, exist_ok=True)           # make sure the screenshot folder exists
    shot = SHOTS / f"{Path(page).parent.name.lower() or 'home'}-{width}.png"
    run_chrome([f"--screenshot={shot}", f"--window-size={width + 40},2400",
                "--force-device-scale-factor=2", wrap.as_uri()], work)

    problems = 0                                       # counts the rules this width breaks
    sideways = m["scrollWidth"] - m["clientWidth"]     # positive means the page scrolls sideways
    if sideways > 0:
        print(f"  {width}px: FAIL page scrolls sideways by {sideways}px "
              f"(scrollWidth {m['scrollWidth']} > clientWidth {m['clientWidth']})")
        problems += 1
    else:
        print(f"  {width}px: ok, no sideways scroll (scrollWidth {m['scrollWidth']} "
              f"<= clientWidth {m['clientWidth']})")

    wide = [c for c in m["charts"] if c["w"] > 521]    # the 520px cap, with a pixel of tolerance
    tiny = [c for c in m["charts"] if c["w"] < 120]    # a chart this narrow is unreadable
    print(f"  {width}px: {len(m['charts'])} charts, widths "
          f"{sorted({c['w'] for c in m['charts']})}, heights {sorted({c['h'] for c in m['charts']})}")
    if wide:
        print(f"  {width}px: FAIL charts wider than 520px: {wide}")
        problems += 1
    if tiny:
        print(f"  {width}px: FAIL charts narrower than 120px: {tiny}")
        problems += 1

    if m["smallTargets"]:
        print(f"  {width}px: FAIL {len(m['smallTargets'])} tap targets under 40px: "
              f"{m['smallTargets'][:5]}")
        problems += 1
    else:
        print(f"  {width}px: ok, every link and button is at least 40px tall")

    scrolling = [s for s in m["scrollers"] if s["scrolls"]]        # the table boxes that do scroll
    unstuck = [s for s in scrolling if s["sticky"] != "sticky"]    # and whose first column is not pinned
    print(f"  {width}px: {len(scrolling)} of {len(m['scrollers'])} table boxes scroll sideways")
    if unstuck:
        print(f"  {width}px: FAIL a scrolling table has no sticky first column: {unstuck}")
        problems += 1

    if m["overflow"]:
        print(f"  {width}px: FAIL elements past the right edge: {m['overflow']}")
        problems += 1

    print(f"  {width}px: screenshot {shot}")
    return problems
    # returns the number of problems found at this width, as an int


def main():
    """Conductor: check the page at each width asked for and report the total."""
    if len(sys.argv) < 2:                              # no page given: say how to call it
        print(__doc__)
        sys.exit(2)
    page = sys.argv[1]                                 # the page to test, relative to the repo root
    widths = [int(w) for w in sys.argv[2:]] or [360, 1200]        # 360 is the phone, 1200 the desktop
    print(f"checking {page}")
    total = 0                                          # problems across every width
    with tempfile.TemporaryDirectory() as tmp:         # one throw-away folder for the whole run
        work = Path(tmp)
        for width in widths:                           # walk the widths in the order given
            total += check(page, width, work)
    print("RESULT:", "ok" if total == 0 else f"{total} problem(s)")
    sys.exit(0 if total == 0 else 1)
    # returns nothing; it prints a report, writes the screenshots and sets the exit status


if __name__ == "__main__":
    main()
