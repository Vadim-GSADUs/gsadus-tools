"""Builds GSADUs.streamDeckProfile (Stream Deck 7.x, MK.2) from the key list below.

Run: python StreamDeck\\build-profile.py, then double-click the .streamDeckProfile to import it.
The format copies an export from the app (FormatVersion 1, profile Version 3.0).
"""
import json
import uuid
import zipfile
from pathlib import Path

HERE = Path(__file__).parent
NS = uuid.UUID("5d1c0a7e-6f2b-4c1e-9a53-3b7f0e2d9c41")  # stable ids, so rebuilds diff cleanly
PROFILE_ID = str(uuid.uuid5(NS, "profile")).upper()
PAGES = ["Dev", "Revit", "Office"]


def website(url):
    return "com.elgato.streamdeck.system.website", "Website", {"browser": "", "openInBrowser": True, "path": url}


def open_(path):
    return "com.elgato.streamdeck.system.open", "Open", {"path": f'"{path}"'}


def text(s):
    hotkey = {"KeyModifiers": 0, "QTKeyCode": 33554431, "VKeyCode": -1}
    return "com.elgato.streamdeck.system.text", "Text", {
        "Hotkey": hotkey, "isSendingEnter": False, "isTypingMode": True, "pastedText": s}


def ctrl(letter):
    code = ord(letter.upper())
    empty = {"KeyCmd": False, "KeyCtrl": False, "KeyModifiers": 0, "KeyOption": False,
             "KeyShift": False, "NativeCode": 146, "QTKeyCode": 33554431, "VKeyCode": -1}
    key = {**empty, "KeyCtrl": True, "KeyModifiers": 2, "NativeCode": code, "QTKeyCode": code, "VKeyCode": code}
    return "com.elgato.streamdeck.system.hotkey", "Hotkey", {"Coalesce": True, "Hotkeys": [key, empty, empty, empty]}


def goto(page):
    return "com.elgato.streamdeck.page.goto", "Go to Page", {"PageIndex": PAGES.index(page) + 1}


DECK = "C:\\GSADUs\\Tools\\StreamDeck"  # the main checkout, never a worktree
DRIVE = "G:\\Shared drives"
# page -> {"col,row": (icon, action)}
KEYS = {
    "Dev": {
        "0,0": ("dev-webapp", website("http://localhost:3009")),
        "1,0": ("dev-site", website("http://localhost:3019")),
        "2,0": ("dev-pm", website("http://localhost:3200")),
        "3,0": ("dev-claude", open_(f"{DECK}\\claude.cmd")),
        "0,1": ("dev-docker", open_(f"{DECK}\\start-docker.cmd")),
        "1,1": ("dev-unwip", open_(f"{DECK}\\unwip-all.cmd")),
        "2,1": ("dev-wip", open_(f"{DECK}\\wip-all.cmd")),
        "0,2": ("dev-vercel", website("https://vercel.com/vadim-7430s-projects")),
        "1,2": ("dev-supabase", website("https://supabase.com/dashboard/project/rfguhdxdrzbqcoiulyuv")),
        "3,2": ("nav-revit", goto("Revit")),
        "4,2": ("nav-office", goto("Office")),
    },
    "Revit": {
        "0,0": ("revit-open", open_("C:\\Program Files\\Autodesk\\Revit 2026\\Revit.exe")),
        "1,0": ("revit-vg", text("VG")),
        "2,0": ("revit-thin", text("TL")),
        "3,0": ("revit-hide", text("EH")),
        "4,0": ("revit-isolate", text("HI")),
        "0,1": ("revit-reset", text("HR")),
        "1,1": ("revit-3d", text("3D")),
        "2,1": ("revit-tile", text("WT")),
        "3,1": ("revit-pyrevit", text("RL")),
        "0,2": ("revit-projects", open_(f"{DRIVE}\\GSADUs Projects")),
        "3,2": ("nav-dev", goto("Dev")),
        "4,2": ("nav-office", goto("Office")),
    },
    "Office": {
        "0,0": ("office-mute", ctrl("D")),
        "1,0": ("office-camera", ctrl("E")),
        "2,0": ("office-meet", website("https://meet.google.com")),
        "0,1": ("office-pipedrive", website("https://goldenstateadus.pipedrive.com")),
        "1,1": ("office-pm", website("https://gsadus-pm.vercel.app")),
        "2,1": ("office-projects", open_(f"{DRIVE}\\GSADUs Projects")),
        "3,1": ("office-business", open_(f"{DRIVE}\\GSADUs Business")),
        "3,2": ("nav-dev", goto("Dev")),
        "4,2": ("nav-revit", goto("Revit")),
    },
}


def page_id(name):
    return str(uuid.uuid5(NS, f"page:{name}")).upper()


def dumps(obj):
    return json.dumps(obj, separators=(",", ":"))


def build(out):
    root = f"Profiles/{PROFILE_ID}.sdProfile"
    plugins = set()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for page, keys in KEYS.items():
            pid = page_id(page)
            actions = {}
            for pos, (icon, (action_uuid, name, settings)) in keys.items():
                plugin = action_uuid.rsplit(".", 1)[0] if action_uuid.startswith("com.elgato.streamdeck.page") else action_uuid
                plugins.add(plugin)
                actions[pos] = {
                    "ActionID": str(uuid.uuid5(NS, f"{page}:{pos}")),
                    "LinkedTitle": True,
                    "Name": name,
                    "Resources": None,
                    "Settings": settings,
                    "State": 0,
                    "States": [{"Image": f"Images/{icon}.png", "ShowTitle": False, "Title": ""}],
                    "UUID": action_uuid,
                }
                z.write(HERE / "icons" / f"{icon}.png", f"{root}/Profiles/{pid}/Images/{icon}.png")
            z.writestr(f"{root}/Profiles/{pid}/manifest.json",
                       dumps({"Controllers": [{"Actions": actions, "Type": "Keypad"}], "Icon": "", "Name": ""}))
        pinned = page_id("pinned")
        z.writestr(f"{root}/Profiles/{pinned}/manifest.json",
                   dumps({"Controllers": [{"Actions": None, "Type": "Keypad"}], "Icon": "", "Name": ""}))
        z.writestr(f"{root}/manifest.json", dumps({
            "AppIdentifier": "*",
            "Device": {"Model": "20GBA9901", "UUID": "d8f5efcf-eb45-4053-b958-5b87934948ad"},
            "Name": "GSADUs",
            "Pages": {"Current": "00000000-0000-0000-0000-000000000000", "Default": pinned.lower(),
                      "Pages": [page_id(p).lower() for p in PAGES]},
            "Version": "3.0",
        }))
        z.writestr("package.json", dumps({
            "AppVersion": "7.6.0.23012", "DeviceModel": "20GBA9901", "DeviceSettings": None,
            "FormatVersion": 1, "OSType": "Windows", "OSVersion": "10.0.26200",
            "RequiredPlugins": sorted(plugins),
        }))


if __name__ == "__main__":
    build(HERE / "GSADUs.streamDeckProfile")
    print(f"wrote {HERE / 'GSADUs.streamDeckProfile'}")
