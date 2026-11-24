"""Keep one live copy of the workbook on OneDrive, shared view-only with anyone who has the link.

  python -m tracker.onedrive connect   # one time: prints a code; sign in at microsoft.com/devicelogin
  python -m tracker.onedrive pull      # download the OneDrive copy and pull your edits/decisions back
  python -m tracker.onedrive push      # upload the rebuilt workbook (same file, same share link)

Secrets (GitHub > Settings > Secrets and variables > Actions):
  MS_CLIENT_ID      Application (client) ID of your free Microsoft app registration
  MS_REFRESH_TOKEN  written automatically by 'connect' (and refreshed on every run)
  GH_PAT            a fine-grained GitHub token with 'Secrets: read and write' on this repo,
                    so the workflow can save the refreshed Microsoft token
The share link is stored in data/onedrive.json and shown on the workbook's Read me tab.
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time

import requests

from .common import DATA, OUTPUT, WORKBOOK_NAME, load_config, load_env, log

AUTH = "https://login.microsoftonline.com/common/oauth2/v2.0"
GRAPH = "https://graph.microsoft.com/v1.0/me/drive"
SCOPE = "Files.ReadWrite offline_access"
INFO = DATA / "onedrive.json"


def _cfg():
    return load_config().get("onedrive", {"folder": "Ethos Intelligence Tracker"})


def _path():
    return f"/root:/{_cfg()['folder']}/{WORKBOOK_NAME}:"


def _save_secret(name: str, value: str) -> bool:
    if not os.environ.get("GH_PAT"):
        print(f"[onedrive] GH_PAT not set: cannot save {name}", flush=True)
        return False
    env = {**os.environ, "GH_TOKEN": os.environ["GH_PAT"]}
    r = subprocess.run(["gh", "secret", "set", name], input=value, text=True, capture_output=True, env=env)
    if r.returncode != 0:
        print(f"[onedrive] could not save {name}: {r.stderr[:200]}", flush=True)
    return r.returncode == 0


def token() -> str:
    load_env()
    cid, rt = os.environ.get("MS_CLIENT_ID", ""), os.environ.get("MS_REFRESH_TOKEN", "")
    if not cid or not rt:
        raise RuntimeError("OneDrive not connected (MS_CLIENT_ID / MS_REFRESH_TOKEN missing)")
    r = requests.post(f"{AUTH}/token", data={"client_id": cid, "grant_type": "refresh_token",
                                              "refresh_token": rt, "scope": SCOPE}, timeout=30)
    js = r.json()
    if "access_token" not in js:
        raise RuntimeError(f"OneDrive sign-in expired: {js.get('error_description', '')[:160]}")
    if js.get("refresh_token") and js["refresh_token"] != rt:
        _save_secret("MS_REFRESH_TOKEN", js["refresh_token"])  # Microsoft rotates it
    return js["access_token"]


def connect() -> None:
    cid = os.environ.get("MS_CLIENT_ID", "")
    r = requests.post(f"{AUTH}/devicecode", data={"client_id": cid, "scope": SCOPE}, timeout=30).json()
    if "user_code" not in r:
        sys.exit(f"Could not start sign-in: {r}")
    print("\n" + "=" * 70 + f"\n  OPEN {r['verification_uri']}  AND ENTER THE CODE:  {r['user_code']}\n"
          "  then sign in with the Microsoft account whose OneDrive should hold the tracker.\n"
          + "=" * 70 + "\n", flush=True)
    deadline = time.time() + int(r.get("expires_in", 900))
    while time.time() < deadline:
        time.sleep(int(r.get("interval", 5)))
        t = requests.post(f"{AUTH}/token", data={"client_id": cid, "device_code": r["device_code"],
                          "grant_type": "urn:ietf:params:oauth:grant-type:device_code"}, timeout=30).json()
        if "refresh_token" in t:
            os.environ["MS_REFRESH_TOKEN"] = t["refresh_token"]
            ok = _save_secret("MS_REFRESH_TOKEN", t["refresh_token"])
            print("[onedrive] connected" + ("" if ok else " (but token NOT saved: add GH_PAT)"), flush=True)
            return
        if t.get("error") not in ("authorization_pending", "slow_down"):
            sys.exit(f"Sign-in failed: {t.get('error_description', t)}")
    sys.exit("Sign-in timed out; run the workflow again.")


BASELINE_ASSET = "onedrive-baseline.xlsx"


def _baseline_upload(f) -> None:
    """Keep a copy of exactly what was uploaded (a release asset), to spot layout edits later."""
    import shutil
    if not os.environ.get("GH_TOKEN"):
        return
    b = OUTPUT / BASELINE_ASSET
    shutil.copy(f, b)
    subprocess.run(["gh", "release", "upload", "latest", str(b), "--clobber"], capture_output=True, text=True)


def _baseline_download():
    if not os.environ.get("GH_TOKEN"):
        return None
    d = OUTPUT / "baseline"
    d.mkdir(parents=True, exist_ok=True)
    for name in (BASELINE_ASSET, WORKBOOK_NAME):  # fall back to the last published workbook
        r = subprocess.run(["gh", "release", "download", "latest", "-p", name, "-D", str(d), "--clobber"],
                           capture_output=True, text=True)
        p = d / name
        if r.returncode == 0 and p.exists():
            return p
    return None


def pull() -> bool:
    """Download the OneDrive copy and pull edits + review decisions back into data/."""
    try:
        h = {"Authorization": f"Bearer {token()}"}
    except RuntimeError as e:
        log("onedrive", "skipped", str(e))
        return False
    meta = requests.get(f"{GRAPH}{_path()}", headers=h, timeout=60)
    if meta.status_code == 404:
        log("onedrive", "ok", "no OneDrive copy yet (first upload pending)")
        return False
    info = json.loads(INFO.read_text()) if INFO.exists() else {}
    if not info.get("uploaded_ctag") or meta.json().get("cTag") == info.get("uploaded_ctag"):
        log("onedrive", "ok", "OneDrive copy unchanged since the last upload; nothing to pull")
        return False
    r = requests.get(f"{GRAPH}{_path()}/content", headers=h, timeout=120)
    if r.status_code == 404:
        log("onedrive", "ok", "no OneDrive copy yet (first upload pending)")
        return False
    r.raise_for_status()
    tmp = OUTPUT / "onedrive_copy.xlsx"
    OUTPUT.mkdir(parents=True, exist_ok=True)
    tmp.write_bytes(r.content)
    from . import sync_from_excel
    n = sync_from_excel.run(str(tmp))
    base = _baseline_download()
    if base:  # keep the owner's layout changes (widths, formatting, formula fixes)
        from .owner_layout import capture
        k = capture(base, tmp)
        log("owner_layout", "ok", f"{k} layout changes kept from the OneDrive copy")
    log("onedrive", "ok", f"pulled edits from OneDrive copy ({n} changes taken)")
    return True


def push() -> str:
    try:
        h = {"Authorization": f"Bearer {token()}"}
    except RuntimeError as e:
        log("onedrive", "skipped", str(e))
        return ""
    f = OUTPUT / WORKBOOK_NAME
    for attempt in range(11):  # about 10 minutes while someone has it open
        r = requests.put(f"{GRAPH}{_path()}/content", headers={**h, "Content-Type":
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
                         data=f.read_bytes(), timeout=180)
        if r.status_code in (200, 201):
            break
        if r.status_code == 423:  # open in Excel: try again shortly
            time.sleep(60)
            continue
        r.raise_for_status()
    else:
        log("onedrive", "failed", "file locked (open in Excel); will retry next run")
        return ""
    item = r.json()
    info = json.loads(INFO.read_text()) if INFO.exists() else {}
    if info.get("item_id") != item["id"] or not info.get("view_link"):
        lk = requests.post(f"{GRAPH}/items/{item['id']}/createLink", headers=h,
                           json={"type": "view", "scope": "anonymous"}, timeout=60).json()
        info = {"item_id": item["id"], "view_link": lk.get("link", {}).get("webUrl", ""),
                "owner_link": item.get("webUrl", "")}
    info["uploaded_ctag"] = item.get("cTag", "")
    from .sync_from_excel import remember_published
    remember_published(f)  # what people will see = the baseline for spotting their edits
    _baseline_upload(f)
    INFO.write_text(json.dumps(info, indent=2))
    log("onedrive", "ok", f"uploaded; view link {info.get('view_link', '')}")
    return info.get("view_link", "")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("action", choices=["connect", "pull", "push"])
    ap.add_argument("--then-push", action="store_true", help="after connecting, upload the workbook")
    a = ap.parse_args()
    {"connect": connect, "pull": pull, "push": push}[a.action]()
    if a.action == "connect" and a.then_push:
        link = push()
        print(f"\n[onedrive] Share this link (view-only, anyone with the link): {link}\n", flush=True)
