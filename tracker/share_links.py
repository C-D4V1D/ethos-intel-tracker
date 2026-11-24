"""Make a view-only, no-sign-in link for every PDF in the tracker's OneDrive folder.

The links are written to "share-links.txt" in that same OneDrive folder (not to the
public run log), so only the owner sees them before choosing to send them.
"""
from __future__ import annotations

import requests

from .common import log
from .onedrive import GRAPH, _cfg, token


def run() -> None:
    h = {"Authorization": f"Bearer {token()}"}
    folder = _cfg()["folder"]
    items = requests.get(f"{GRAPH}/root:/{folder}:/children", headers=h, timeout=60).json().get("value", [])
    lines = []
    for it in items:
        if not it.get("name", "").lower().endswith(".pdf"):
            continue
        lk = requests.post(f"{GRAPH}/items/{it['id']}/createLink", headers=h,
                           json={"type": "view", "scope": "anonymous"}, timeout=60).json()
        lines.append(f"{it['name']}\t{lk.get('link', {}).get('webUrl', 'FAILED')}")
    requests.put(f"{GRAPH}/root:/{folder}/share-links.txt:/content",
                 headers={**h, "Content-Type": "text/plain"}, data="\n".join(lines).encode(), timeout=60)
    log("share_links", "ok", f"{len(lines)} view-only links written to OneDrive/{folder}/share-links.txt")
    print(f"{len(lines)} links written to OneDrive (not shown here)")


if __name__ == "__main__":
    run()
