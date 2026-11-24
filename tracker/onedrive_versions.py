"""Download recent versions of the OneDrive copy (for recovering the owner's edits).

Saves each version to output/versions/<timestamp>.xlsx and prints the list
(time and size only).
"""
from __future__ import annotations

import requests

from .common import OUTPUT
from .onedrive import GRAPH, _path, token


def run(limit: int = 12) -> None:
    h = {"Authorization": f"Bearer {token()}"}
    vs = requests.get(f"{GRAPH}{_path()}/versions", headers=h, timeout=60).json().get("value", [])
    out = OUTPUT / "versions"
    out.mkdir(parents=True, exist_ok=True)
    for v in vs[:limit]:
        stamp = v["lastModifiedDateTime"].replace(":", "-")
        who = (v.get("lastModifiedBy", {}).get("user", {}) or {}).get("displayName", "")
        r = requests.get(f"{GRAPH}{_path()}/versions/{v['id']}/content", headers=h, timeout=120)
        if r.ok:
            (out / f"{stamp}.xlsx").write_bytes(r.content)
        print(stamp, v.get("size"), who, "saved" if r.ok else f"HTTP {r.status_code}")


if __name__ == "__main__":
    run()
