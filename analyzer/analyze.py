"""
Mail analyser - runs on GitHub Actions (or locally), standard library only.
Reads the text feed pushed by collector/ and analyses new mails with an LLM
through OpenRouter.

Input:   feed/YYYY-MM-DD/mails/0001.txt
Output:  reports/YYYY-MM-DD.json   per-mail analysis (cache: analysed mails are skipped)
         reports/YYYY-MM-DD.md     daily report: digest + mail table
         DASHBOARD.md              actions, deadlines and pending replies of recent days

Env:     OPENROUTER_API_KEY (required), OPENROUTER_MODEL, MAIL_OWNER
Usage:   python analyzer/analyze.py [--day YYYY-MM-DD] [--force]
"""
import os, re, sys, json, glob, time, argparse, datetime as dt, urllib.request, urllib.error

# ---------------- settings ----------------
DEFAULT_MODEL = "google/gemini-2.5-flash"   # cheap, long context; override with OPENROUTER_MODEL
MAX_MAIL_CHARS = 20000           # chars of one mail sent to the model
BATCH_CHARS = 60000              # chars per request (several mails per request)
DASHBOARD_DAYS = 14              # how many days DASHBOARD.md looks back
# -----------------------------------------

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.environ.get("MAIL_ANALYSER_ROOT") or os.path.dirname(HERE)
FEED = os.path.join(ROOT, "feed")
REPORTS = os.path.join(ROOT, "reports")
API_URL = os.environ.get("OPENROUTER_BASE_URL", "https://openrouter.ai/api/v1").rstrip("/") + "/chat/completions"
MODEL = os.environ.get("OPENROUTER_MODEL") or DEFAULT_MODEL
OWNER = os.environ.get("MAIL_OWNER") or "mail sahibi"
PRIO = {"yüksek": 0, "orta": 1, "düşük": 2}
HEADER = re.compile(r"^(No|Date|Folder|From|To|CC|Subject): ?(.*)$", re.M)

cost = {"usd": 0.0, "tokens": 0, "calls": 0}


def prompt(name):
    with open(os.path.join(HERE, "prompts", name), encoding="utf-8") as f:
        return f.read().replace("{owner}", OWNER)


# ---------- OpenRouter ----------
def llm(system, user, json_mode):
    body = {"model": MODEL, "temperature": 0.2, "usage": {"include": True},
            "messages": [{"role": "system", "content": system}, {"role": "user", "content": user}]}
    if json_mode:
        body["response_format"] = {"type": "json_object"}
    req = urllib.request.Request(API_URL, data=json.dumps(body).encode("utf-8"), headers={
        "Authorization": "Bearer " + os.environ["OPENROUTER_API_KEY"],
        "Content-Type": "application/json",
        "X-Title": "Mail-Analyser"})
    err = None
    for attempt in range(4):
        try:
            with urllib.request.urlopen(req, timeout=300) as r:
                data = json.load(r)
            if "error" in data:
                raise RuntimeError(f"OpenRouter: {data['error']}")
            usage = data.get("usage") or {}
            cost["usd"] += usage.get("cost") or 0
            cost["tokens"] += usage.get("total_tokens") or 0
            cost["calls"] += 1
            return data["choices"][0]["message"]["content"] or ""
        except urllib.error.HTTPError as e:
            msg = e.read().decode("utf-8", "ignore")[:400]
            if e.code in (400, 401, 402, 403, 404):    # bad key / no credit / bad model: retry won't help
                raise RuntimeError(f"OpenRouter HTTP {e.code}: {msg}")
            err = f"HTTP {e.code}: {msg}"
        except (urllib.error.URLError, TimeoutError, KeyError, IndexError, RuntimeError) as e:
            err = repr(e)
        time.sleep(5 * 2 ** attempt)
    raise RuntimeError(f"OpenRouter failed after retries: {err}")


def parse_json(text):
    text = re.sub(r"^```(json)?|```$", "", text.strip(), flags=re.M)
    return json.loads(text[text.index("{"): text.rindex("}") + 1])


# ---------- feed ----------
def read_mail(path):
    with open(path, encoding="utf-8") as f:
        text = f.read()
    head = {}
    for k, v in HEADER.findall(text[:3000]):
        head.setdefault(k, v.strip())   # first match only: quoted replies repeat From:/To:
    return {"no": head.get("No") or os.path.basename(path)[:-4],
            "date": head.get("Date", ""),
            "dir": (head.get("Folder", "").split("/")[0].strip() or "IN"),
            "from": head.get("From", ""),
            "to": head.get("To", ""),
            "subject": head.get("Subject", ""),
            "text": text[:MAX_MAIL_CHARS] + ("\n... [cut]" if len(text) > MAX_MAIL_CHARS else "")}


def batches(mails):
    cur, size = [], 0
    for m in mails:
        if cur and size + len(m["text"]) > BATCH_CHARS:
            yield cur
            cur, size = [], 0
        cur.append(m)
        size += len(m["text"])
    if cur:
        yield cur


# ---------- analysis ----------
def analyse_batch(mails):
    user = "\n\n".join(f"===== MAIL {m['no']} =====\n{m['text']}" for m in mails)
    out = parse_json(llm(prompt("mail.md"), user, json_mode=True))
    res = {}
    for a in out.get("mails", []):
        no = str(a.get("no", "")).strip().zfill(4)
        if any(m["no"] == no for m in mails):
            res[no] = a
    return res


def digest(day, rep):
    rows = [{k: v for k, v in m.items() if k not in ("to",)} for m in rep["mails"].values()]
    user = f"Tarih: {day}\n\n" + json.dumps(rows, ensure_ascii=False, indent=1)
    return llm(prompt("day.md"), user, json_mode=False).strip()


def process_day(day, force):
    files = sorted(glob.glob(os.path.join(FEED, day, "mails", "*.txt")))
    path = os.path.join(REPORTS, day + ".json")
    rep = {"day": day, "mails": {}, "digest": ""}
    if os.path.exists(path) and not force:
        with open(path, encoding="utf-8") as f:
            rep = json.load(f)
    new = [m for m in map(read_mail, files) if m["no"] not in rep["mails"]]
    if not new and rep["digest"]:
        return 0, 0
    failed = 0
    for b in batches(new):
        try:
            res = analyse_batch(b)
        except Exception as e:
            print(f"  {day}: batch {b[0]['no']}-{b[-1]['no']} failed: {e}")
            failed += len(b)
            continue
        for m in b:
            if m["no"] not in res:
                failed += 1     # model skipped it; retried next run
                continue
            meta = {k: m[k] for k in ("no", "date", "dir", "from", "to", "subject")}
            rep["mails"][m["no"]] = {**meta, **res[m["no"]], "no": m["no"]}
    if rep["mails"]:
        rep["model"] = MODEL
        try:
            rep["digest"] = digest(day, rep)
        except Exception as e:
            print(f"  {day}: digest failed: {e}")
            rep["digest"] = ""      # empty digest -> retried next run
            failed += 1
    os.makedirs(REPORTS, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(rep, f, ensure_ascii=False, indent=1)
    write_day_md(rep)
    return len(new) - failed, failed


# ---------- markdown ----------
def cell(s, n=160):
    s = re.sub(r"\s+", " ", str(s or "")).replace("|", "/")
    return s[:n] + "…" if len(s) > n else s


def sort_key(m):
    return (PRIO.get(m.get("priority"), 3), m.get("deadline") or "9999", m["no"])


def write_day_md(rep):
    mails = sorted(rep["mails"].values(), key=sort_key)
    lines = [f"# Mail raporu — {rep['day']}", "",
             f"[← Dashboard](../DASHBOARD.md) · {len(mails)} mail · model: `{rep.get('model', '-')}`", "",
             rep["digest"], "", "## Tüm mailler", "",
             "| No | Tarih | Yön | Kimden | Konu | Öncelik | Özet | Aksiyon |",
             "|---|---|---|---|---|---|---|---|"]
    for m in mails:
        lines.append(f"| {m['no']} | {cell(m.get('date'))} | {m.get('dir')} | {cell(m.get('from'), 50)} | "
                     f"{cell(m.get('subject'), 80)} | {m.get('priority', '')} | {cell(m.get('summary'))} | "
                     f"{cell(m.get('action') if m.get('action_needed') else '')} |")
    with open(os.path.join(REPORTS, rep["day"] + ".md"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def write_dashboard():
    days = sorted(os.path.basename(p)[:-5] for p in glob.glob(os.path.join(REPORTS, "*.json")))
    recent = days[-DASHBOARD_DAYS:]
    actions, deadlines, waiting = [], [], []
    for d in recent:
        with open(os.path.join(REPORTS, d + ".json"), encoding="utf-8") as f:
            for m in json.load(f)["mails"].values():
                m["day"] = d
                if m.get("action_needed"):
                    actions.append(m)
                if m.get("deadline"):
                    deadlines.append(m)
                if m.get("waiting_on"):
                    waiting.append(m)
    today = dt.date.today().isoformat()
    ref = lambda m: f"[{m['day']} #{m['no']}](reports/{m['day']}.md)"
    out = [f"# Mail Dashboard", "",
           f"Güncellendi: {dt.datetime.now(dt.timezone.utc):%Y-%m-%d %H:%M} UTC · son {len(recent)} gün", ""]
    if days:
        last = days[-1]
        with open(os.path.join(REPORTS, last + ".json"), encoding="utf-8") as f:
            out += [f"## Son gün: [{last}](reports/{last}.md)", "", json.load(f).get("digest", ""), ""]
    out += ["## Yapılacaklar", "", "| Öncelik | Aksiyon | Son tarih | Konu | Mail |", "|---|---|---|---|---|"]
    out += [f"| {m.get('priority', '')} | {cell(m.get('action'))} | {m.get('deadline') or ''} | "
            f"{cell(m.get('subject'), 70)} | {ref(m)} |" for m in sorted(actions, key=sort_key)]
    out += ["", "## Son tarihler", "", "| Tarih | Ne | Konu | Mail |", "|---|---|---|---|"]
    for m in sorted(deadlines, key=lambda m: m["deadline"]):
        flag = " ⚠️" if m["deadline"] < today else ""
        out.append(f"| {m['deadline']}{flag} | {cell(m.get('action') or m.get('summary'))} | "
                   f"{cell(m.get('subject'), 70)} | {ref(m)} |")
    out += ["", "## Cevap beklediklerim", "", "| Kimden / ne | Konu | Mail |", "|---|---|---|"]
    out += [f"| {cell(m['waiting_on'])} | {cell(m.get('subject'), 70)} | {ref(m)} |" for m in waiting]
    out += ["", "## Günlük raporlar", ""] + [f"- [{d}](reports/{d}.md)" for d in reversed(days)]
    with open(os.path.join(ROOT, "DASHBOARD.md"), "w", encoding="utf-8") as f:
        f.write("\n".join(out) + "\n")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--day", help="only this feed day (YYYY-MM-DD)")
    ap.add_argument("--force", action="store_true", help="re-analyse mails that are already done")
    a = ap.parse_args()
    if not os.environ.get("OPENROUTER_API_KEY"):
        sys.exit("OPENROUTER_API_KEY is not set (GitHub: Settings > Secrets and variables > Actions)")
    days = [a.day] if a.day else sorted(os.path.basename(p) for p in glob.glob(os.path.join(FEED, "*")) if os.path.isdir(p))
    total_fail = 0
    for d in days:
        ok, fail = process_day(d, a.force)
        total_fail += fail
        if ok or fail:
            print(f"{d}: {ok} analysed, {fail} failed")
    write_dashboard()
    print(f"model {MODEL} · {cost['calls']} calls · {cost['tokens']} tokens · ${cost['usd']:.4f}")
    if total_fail:
        sys.exit(f"{total_fail} mails failed - they will be retried on the next run")


if __name__ == "__main__":
    main()
