"""Builds GSADUs.streamDeckProfile (Stream Deck 7.x, MK.2) from the key list below.

Run: python StreamDeck\\build-profile.py, then import the .streamDeckProfile (README).
The format copies an export from the app (FormatVersion 1, profile Version 3.0).
"""
import json
import uuid
import zipfile
from pathlib import Path

HERE = Path(__file__).parent
NS = uuid.UUID("5d1c0a7e-6f2b-4c1e-9a53-3b7f0e2d9c41")  # stable ids, so rebuilds diff cleanly
PROFILE_ID = str(uuid.uuid5(NS, "profile")).upper()


def website(url):
    return "com.elgato.streamdeck.system.website", "Website", {"browser": "", "openInBrowser": True, "path": url}


def open_(path):
    return "com.elgato.streamdeck.system.open", "Open", {"path": f'"{path}"'}


DECK = "C:\\GSADUs\\Tools\\StreamDeck"  # the main checkout, never a worktree
DRIVE = "G:\\Shared drives"
ME = "authuser=vadim@gsadus.com"
# "col,row": (icon, action)
KEYS = {
    "0,0": ("gmail", website(f"https://mail.google.com/mail/?{ME}")),
    "1,0": ("chat", website(f"https://chat.google.com/?{ME}")),
    "2,0": ("pm", website("https://gsadus-pm.vercel.app")),
    "3,0": ("webapp", website("https://gsadus.vercel.app")),
    "4,0": ("site", website("https://gsadus.com")),
    "0,1": ("pipedrive", website("https://goldenstateadus.pipedrive.com")),
    "1,1": ("proposals", open_(f"{DRIVE}\\GSADUs Business\\1 - Proposals")),
    "2,1": ("projects", open_(f"{DRIVE}\\GSADUs Projects")),
    "3,1": ("helpdesk", website("https://gsadus-helpdesk.vercel.app")),
    "4,1": ("board", open_(f"{DECK}\\board.cmd")),
    "0,2": ("orchestrator", open_(f"{DECK}\\claude.cmd")),
    "1,2": ("gws-login", open_(f"{DECK}\\gws-login.cmd")),
    "2,2": ("gcloud-login", open_(f"{DECK}\\gcloud-login.cmd")),
    "3,2": ("remote", open_(f"{DECK}\\remote.cmd")),
    "4,2": ("stop-servers", open_(f"{DECK}\\stop-dev-servers.cmd")),
}


def dumps(obj):
    return json.dumps(obj, separators=(",", ":"))


def build(out):
    root = f"Profiles/{PROFILE_ID}.sdProfile"
    page = str(uuid.uuid5(NS, "page:main")).upper()
    pinned = str(uuid.uuid5(NS, "page:pinned")).upper()
    actions = {}
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for pos, (icon, (action_uuid, name, settings)) in KEYS.items():
            actions[pos] = {
                "ActionID": str(uuid.uuid5(NS, f"main:{pos}")),
                "LinkedTitle": True,
                "Name": name,
                "Resources": None,
                "Settings": settings,
                "State": 0,
                "States": [{"Image": f"Images/{icon}.png", "ShowTitle": False, "Title": ""}],
                "UUID": action_uuid,
            }
            z.write(HERE / "icons" / f"{icon}.png", f"{root}/Profiles/{page}/Images/{icon}.png")
        z.writestr(f"{root}/Profiles/{page}/manifest.json",
                   dumps({"Controllers": [{"Actions": actions, "Type": "Keypad"}], "Icon": "", "Name": ""}))
        z.writestr(f"{root}/Profiles/{pinned}/manifest.json",
                   dumps({"Controllers": [{"Actions": None, "Type": "Keypad"}], "Icon": "", "Name": ""}))
        z.writestr(f"{root}/manifest.json", dumps({
            "AppIdentifier": "*",
            "Device": {"Model": "20GBA9901", "UUID": "d8f5efcf-eb45-4053-b958-5b87934948ad"},
            "Name": "GSADUs",
            "Pages": {"Current": "00000000-0000-0000-0000-000000000000", "Default": pinned.lower(),
                      "Pages": [page.lower()]},
            "Version": "3.0",
        }))
        z.writestr("package.json", dumps({
            "AppVersion": "7.6.0.23012", "DeviceModel": "20GBA9901", "DeviceSettings": None,
            "FormatVersion": 1, "OSType": "Windows", "OSVersion": "10.0.26200",
            "RequiredPlugins": sorted({a["UUID"] for a in actions.values()}),
        }))


if __name__ == "__main__":
    build(HERE / "GSADUs.streamDeckProfile")
    print(f"wrote {HERE / 'GSADUs.streamDeckProfile'}")
