# -*- coding: utf-8 -*-
"""Parse a candidate's CV (docx / txt / pdf) into a structured profile JSON.

Best-effort sectioning for typical reverse-chronological CVs. Always keeps the
full `raw_text` so downstream tailoring can reason over the original.

Usage:
  python3 parse_cv.py --in cv.docx --out profile.json
  python3 parse_cv.py --in cv.txt  --out profile.json
"""
import sys, re, json, os
from pathlib import Path


def read_docx(path):
    try:
        from docx import Document
    except Exception as e:
        return None, f"python-docx missing: {e}"
    try:
        d = Document(path)
        paras = [p.text for p in d.paragraphs]
        for t in d.tables:
            for row in t.rows:
                for cell in row.cells:
                    paras.append(cell.text)
        return "\n".join(paras), None
    except Exception as e:
        return None, f"docx error: {e}"


def read_txt(path):
    try:
        return Path(path).read_text(encoding="utf-8", errors="ignore"), None
    except Exception as e:
        return None, str(e)


def read_pdf(path):
    try:
        from pdfminer.high_level import extract_text
        return extract_text(str(path)), None
    except Exception:
        pass
    import subprocess
    try:
        out = subprocess.run(["pdftotext", str(path), "-"], capture_output=True, text=True)
        if out.returncode == 0:
            return out.stdout, None
    except Exception:
        pass
    return None, "pdf unavailable (install pdfminer.six or poppler)"


def section(text, start_kws, end_kws):
    """Return the slice of text between the first start_kw and the first end_kw."""
    low = text.lower()
    s = None
    for k in start_kws:
        i = low.find(k)
        if i != -1 and (s is None or i < s):
            s = i
    if s is None:
        return ""
    e = len(text)
    for k in end_kws:
        i = low.find(k, s + 5)
        if i != -1 and i < e:
            e = i
    return text[s:e]


def parse_profile(text):
    prof = {"name": "", "tagline": "", "email": "", "phone": "",
            "summary": [], "experience": [], "education": [], "raw_text": text or ""}
    if not text:
        return prof
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    m = re.search(r"[\w.+-]+@[\w-]+\.[\w.-]+", text)
    if m:
        prof["email"] = m.group(0)

    m = re.search(r"(?:\+?\d[\d\s\-().]{7,}\d)", text)
    if m:
        ph = re.sub(r"[^\d+]+", " ", m.group(0)).strip()
        if len(re.sub(r"\D", "", ph)) >= 8:
            prof["phone"] = ph

    # name: a short title-case / all-caps line near the top, mostly letters
    for l in lines[:8]:
        if (len(l) <= 40 and not re.search(r"@", l)
                and not re.match(r"[\d\-]{7,}", l)
                and sum(c.isalpha() for c in l) >= len(l) * 0.5):
            prof["name"] = l
            break

    # tagline: a line right after name that looks like a role summary
    if prof["name"]:
        try:
            idx = lines.index(prof["name"])
            for l in lines[idx + 1: idx + 3]:
                if 15 <= len(l) <= 120 and not re.search(r"@|http|\d{3,}", l):
                    prof["tagline"] = l
                    break
        except ValueError:
            pass

    # summary / experience / education blocks
    sum_block = section(text, ["professional summary", "summary", "profile", "about me"],
                        ["experience", "employment", "work history", "education"])
    exp_block = section(text, ["experience", "employment", "work history", "professional experience"],
                       ["education", "academic", "skills", "languages"])
    edu_block = section(text, ["education", "academic background", "academic"],
                       ["skills", "languages", "interests", "references"])

    if sum_block:
        prof["summary"] = [b.strip(" •-\t") for b in sum_block.splitlines()
                           if b.strip() and len(b.strip()) > 12][:6]

    prof["experience"] = parse_entries(exp_block)
    prof["education"] = parse_entries(edu_block, education=True)
    return prof


def parse_entries(block, education=False):
    """Crude entry extraction: lines that look like headers + bullet children."""
    entries = []
    if not block:
        return entries
    cur = None
    for ln in block.splitlines():
        s = ln.strip()
        if not s:
            continue
        is_bullet = bool(re.match(r"^[\-\u2022•\*\d][.)]?\s+", s))
        if not is_bullet and (len(s) <= 80) and not re.match(r"^\d{4}", s):
            if cur is None or len(s) < 60:
                if cur:
                    entries.append(cur)
                cur = {"header": s, "bullets": []}
            elif cur:
                cur["bullets"].append(re.sub(r"^[\-\u2022•\*\d][.)]?\s+", "", s))
        else:
            if cur is None:
                cur = {"header": "(untitled)", "bullets": []}
            cur["bullets"].append(re.sub(r"^[\-\u2022•\*\d][.)]?\s+", "", s))
    if cur:
        entries.append(cur)
    # split header into company/role/period where possible
    out = []
    for e in entries:
        h = e["header"]
        period = ""
        m = re.search(r"((?:19|20)\d{2}\s*[-–]\s*(?:19|20)\d{2}|present|current|至今|now)", h, re.I)
        if m:
            period = m.group(0)
        out.append({"company": h, "role": "", "period": period, "bullets": e["bullets"]})
    return out


def main():
    ap = __import__("argparse").ArgumentParser()
    ap.add_argument("--in", dest="src", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    p = Path(args.src)
    if not p.exists():
        print(f"找不到 CV 文件: {p}")
        sys.exit(1)
    ext = p.suffix.lower()
    if ext == ".docx":
        text, err = read_docx(p)
    elif ext == ".pdf":
        text, err = read_pdf(p)
    else:
        text, err = read_txt(p)
    if text is None:
        print(f"读取失败: {err}")
        sys.exit(1)

    prof = parse_profile(text)
    Path(args.out).write_text(json.dumps(prof, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已解析 CV -> {args.out}")
    print(f"  name={prof['name']!r} email={prof['email']!r} phone={prof['phone']!r}")
    print(f"  summary 条数={len(prof['summary'])} | experience 条数={len(prof['experience'])} | education 条数={len(prof['education'])}")
    print(f"  raw_text 长度={len(prof['raw_text'])}")


if __name__ == "__main__":
    main()
