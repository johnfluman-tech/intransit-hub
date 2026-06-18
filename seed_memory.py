"""Seed all memory .md files into the Intransit Hub D1 via /api/memory."""
import os, re, requests

HUB = "https://intransit-hub.intransit-sales.workers.dev"
SEC = "CXEzt4BVqliSmNwdorZf86DuKyceAxjn"
MEM_DIR = r"C:\Users\fluma\.claude\projects\C--Users-fluma\memory"
HEADERS = {"Authorization": f"Bearer {SEC}"}

def post_memory(slug, description, mem_type, body):
    r = requests.post(f"{HUB}/api/memory", json={"slug": slug, "description": description,
        "type": mem_type, "body": body}, headers=HEADERS, timeout=10)
    return r.status_code, r.text

for fname in os.listdir(MEM_DIR):
    if not fname.endswith(".md") or fname == "MEMORY.md":
        continue
    path = os.path.join(MEM_DIR, fname)
    content = open(path, encoding="utf-8").read()

    # Split on --- delimiters
    parts = re.split(r"(?m)^---\s*$", content, maxsplit=2)
    if len(parts) < 3:
        print(f"SKIP {fname} (no frontmatter)")
        continue

    fm, body = parts[1], parts[2].strip()

    slug = re.search(r"^name:\s*(.+)$", fm, re.M)
    slug = slug.group(1).strip() if slug else fname.replace(".md", "")
    desc = re.search(r"^description:\s*(.+)$", fm, re.M)
    desc = desc.group(1).strip() if desc else ""
    mtype = re.search(r"^\s+type:\s*(.+)$", fm, re.M)
    mtype = mtype.group(1).strip() if mtype else "feedback"

    status, resp = post_memory(slug, desc, mtype, body)
    print(f"  {status}  {slug}")

print("\nDone.")
