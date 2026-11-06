#!/usr/bin/env python3
"""Create a five-minute animated Outlook Classic PST training simulation.

Requires: Python 3, Pillow, ffmpeg, and espeak-ng.
The output is a 1280x720 MP4 with animated mouse movement, clicks, menus,
Outlook-style dialogs, typed text, step captions, and English narration.
"""
from pathlib import Path
import math
import shutil
import subprocess
import sys

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "output"
WORK = ROOT / "work"
OUT.mkdir(exist_ok=True)
WORK.mkdir(exist_ok=True)

W, H, FPS = 1280, 720, 10
TOTAL_SECONDS = 300
PST_NAME = "Archive_1126.pst"

# Each scene duration totals exactly 300 seconds.
SCENES = [
    {
        "duration": 20, "title": "OUTLOOK CLASSIC PST ARCHIVE TRAINING",
        "subtitle": "Create • Configure • Test",
        "narration": "Welcome to this Microsoft Outlook Classic training. In this walkthrough, we will create a monthly Personal Folders file, set the data file as the default, configure general rules for incoming and sent messages, and verify the archive. This is an animated demonstration using a sample Outlook profile.",
        "steps": [
            ("Open Outlook Classic", "home", (590, 280), 7),
            ("Review the mailbox and folder pane", "home", (190, 360), 6),
            ("Today: configure a local PST archive", "home", (830, 205), 7),
        ],
    },
    {
        "duration": 50, "title": "1. CREATE A NEW PST FILE",
        "subtitle": "Create the monthly archive data file",
        "narration": "First, open the File tab in Outlook Classic. Select Account Settings, and then choose Account Settings again from the menu. In the Account Settings window, open the Data Files tab. Select Add to create a new Outlook Data File. In the file name field, enter Archive underscore one one two six dot P S T. This follows the agreed naming format, name underscore month and two digit year. Select OK or Save, depending on the dialog shown. The new PST should now appear in the list of Outlook data files. Keep the file in an approved location and use the same naming convention for each archive period.",
        "steps": [
            ("Click File", "file_menu", (78, 85), 7),
            ("Open Account Settings", "account_menu", (145, 220), 8),
            ("Open the Data Files tab", "data_files", (420, 225), 8),
            ("Select Add", "add_dialog", (1045, 490), 8),
            ("Enter Archive_1126.pst", "new_pst", (630, 390), 10),
            ("Save the new Outlook Data File", "data_files_added", (1015, 535), 9),
        ],
    },
    {
        "duration": 35, "title": "2. SET THE DEFAULT DATA FILE",
        "subtitle": "Select the newly created PST",
        "narration": "Next, return to Account Settings and the Data Files tab. Select Archive underscore one one two six dot P S T in the data file list. Choose Set as Default. If Outlook asks for confirmation, review the prompt and confirm the change. Check that the default indicator now appears beside the selected PST. Important: setting a default data file does not move a Microsoft 365 or Exchange mailbox to the PST, and it does not automatically change the server delivery location for new email. It sets the default Outlook data file for applicable local data storage and profile items.",
        "steps": [
            ("Select Archive_1126.pst", "data_files_added", (590, 345), 9),
            ("Click Set as Default", "default_confirm", (1040, 520), 10),
            ("Confirm the default data file", "default_done", (720, 405), 9),
            ("Verify the default indicator", "data_files_default", (585, 345), 7),
        ],
    },
    {
        "duration": 65, "title": "3. CREATE GENERAL INBOX RULES",
        "subtitle": "Apply a rule to messages you receive",
        "narration": "Now configure a general rule for incoming messages. From the Home tab, open Rules and select Manage Rules and Alerts. Choose New Rule. In the Rules Wizard, select the option to apply a rule on messages I receive. Do not select a sender-specific condition. Continue to the actions page and choose the action to move it to the specified folder. Select the folder picker, expand the Archive PST, and choose its Inbox folder. Confirm the destination. Finish the wizard and give the rule a clear name, Inbox Archive Rule. Review the rule summary. This is a general rule and is not restricted to a particular sender. Because the destination is a local PST, this type of rule is client-side and Outlook Classic may need to remain open for it to run. The exact available options can vary by account type and Outlook configuration.",
        "steps": [
            ("Open Rules", "rules_menu", (365, 145), 8),
            ("Choose Manage Rules and Alerts", "rules_alerts", (490, 210), 8),
            ("Select New Rule", "wizard_start", (1080, 565), 8),
            ("Choose: apply rule on messages I receive", "wizard_received", (510, 340), 10),
            ("Continue without sender conditions", "wizard_conditions", (1080, 590), 8),
            ("Choose: move it to the specified folder", "wizard_action", (500, 420), 10),
            ("Select Archive PST > Inbox", "folder_picker_inbox", (650, 430), 8),
            ("Finish: Inbox Archive Rule", "rule_saved_inbox", (1080, 585), 5),
        ],
    },
    {
        "duration": 60, "title": "4. CREATE SENT ITEMS RULES",
        "subtitle": "Apply a general rule to messages you send",
        "narration": "Next, create a separate rule for sent messages. Open Rules and Alerts and select New Rule. Choose the option to apply a rule on messages I send. Do not add a recipient-specific condition. On the actions page, choose the option to move a copy to the specified folder, where that option is available. Select the folder picker, expand Archive underscore one one two six, and choose Sent Items. Confirm the destination and finish the wizard. Name this rule Sent Items Archive Rule. Review the rule summary before saving. The sent-message rule is separate from the incoming-message rule. It is intended to retain a copy of sent messages in the PST archive. Local PST actions are client-side, so keep Outlook Classic running when the rule needs to process messages. Verify the action wording in your Outlook version before applying it to a production mailbox.",
        "steps": [
            ("Open Rules and Alerts", "rules_alerts", (490, 210), 8),
            ("Start a New Rule", "wizard_start", (1080, 565), 7),
            ("Choose: apply rule on messages I send", "wizard_sent", (510, 365), 10),
            ("Choose the sent-message copy action", "wizard_sent_action", (510, 440), 10),
            ("Select Archive PST > Sent Items", "folder_picker_sent", (650, 455), 10),
            ("Finish: Sent Items Archive Rule", "rule_saved_sent", (1080, 585), 8),
            ("Verify both rules are enabled", "rules_list", (520, 330), 7),
        ],
    },
    {
        "duration": 50, "title": "5. TESTING AND VERIFICATION",
        "subtitle": "Check incoming and sent message archiving",
        "narration": "Now test the configuration with sample messages. First, receive a test email in the mailbox. With Outlook running, the general Inbox rule processes the message and moves it to the Inbox folder inside the Archive PST. Next, compose and send a test email. The sent-message rule processes the sent message and retains a copy in the Sent Items folder inside the PST, when the configured action is supported. Open the PST in the folder pane and inspect both destination folders. Confirm the sample incoming and sent messages are present. Also open Rules and Alerts and verify that both rules are enabled. If a message does not appear, check the rule conditions, destination folder, account type, and whether Outlook was running. Test with non-critical sample messages before using the configuration for business email.",
        "steps": [
            ("Receive a sample incoming email", "test_incoming", (850, 345), 12),
            ("Verify it in Archive PST > Inbox", "test_inbox_done", (205, 485), 12),
            ("Send a sample email", "test_sent", (1030, 180), 12),
            ("Verify it in Archive PST > Sent Items", "test_sent_done", (205, 535), 14),
        ],
    },
    {
        "duration": 20, "title": "TRAINING COMPLETE",
        "subtitle": "PST archive configured and checked",
        "narration": "You have completed the Outlook Classic PST archive walkthrough. You created the monthly PST, reviewed the default data file setting, configured general Inbox and Sent Items rules, and checked the results. Back up PST files regularly and follow your organization's email retention policy. Thank you for watching.",
        "steps": [
            ("Review the completed archive", "final", (850, 350), 10),
            ("Thank you for watching", "final", (850, 500), 10),
        ],
    },
]

# Outlook / Windows-like palette
NAVY = "#172b4d"
BLUE = "#0f6cbd"
LIGHT_BLUE = "#e8f2fc"
BORDER = "#c8d1dc"
TEXT = "#243447"
MUTED = "#607080"
WHITE = "#ffffff"
GREEN = "#107c41"

FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
]
BOLD_CANDIDATES = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
]

def load_font(size, bold=False):
    candidates = BOLD_CANDIDATES if bold else FONT_CANDIDATES
    for candidate in candidates:
        if Path(candidate).exists():
            return ImageFont.truetype(candidate, size)
    return ImageFont.load_default()

FONTS = {s: load_font(s) for s in (12, 14, 16, 18, 20, 22, 24, 28, 32)}
BOLDS = {s: load_font(s, True) for s in (12, 14, 16, 18, 20, 22, 24, 28, 32)}

def font(size, bold=False):
    return (BOLDS if bold else FONTS).get(size) or load_font(size, bold)

def rect(d, xy, fill, outline=None, radius=7, width=1):
    d.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=width)

def label(d, xy, text, size=16, fill=TEXT, bold=False):
    d.text(xy, text, font=font(size, bold), fill=fill)

def button(d, xy, text, primary=False, width=None):
    x1, y1, x2, y2 = xy
    fill = BLUE if primary else "#f7f9fb"
    outline = BLUE if primary else "#aeb9c5"
    rect(d, xy, fill, outline, 5, 1)
    f = font(14, primary)
    bbox = d.textbbox((0, 0), text, font=f)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    d.text((x1 + (x2-x1-tw)/2, y1 + (y2-y1-th)/2 - 2), text, font=f, fill=WHITE if primary else TEXT)

def checkbox(d, x, y, text, checked=False, size=15):
    d.rectangle((x, y, x+size, y+size), fill=WHITE, outline="#718096", width=1)
    if checked:
        d.line((x+3,y+8,x+6,y+12,x+13,y+3), fill=BLUE, width=2)
    label(d, (x+size+9,y-2), text, 14)

def draw_outlook_base(d, section):
    # Windows desktop background and taskbar
    d.rectangle((0, 0, W, H), fill="#dce7f2")
    d.ellipse((-230, -300, 500, 430), fill="#c5d9eb")
    d.ellipse((920, 400, 1500, 980), fill="#c7d9e9")
    # Outlook application window
    rect(d, (42, 25, 1238, 666), WHITE, "#8797a8", 10, 2)
    d.rectangle((44, 27, 1236, 65), fill="#f4f6f8")
    label(d, (62, 37), "Outlook", 16, TEXT, True)
    label(d, (1120, 35), "—     □     ×", 16, MUTED)
    # Ribbon tabs and command row
    d.rectangle((44, 66, 1236, 151), fill=WHITE)
    tabs = [("File", 65), ("Home", 115), ("Send / Receive", 185), ("Folder", 320), ("View", 390), ("Help", 450)]
    for name, x in tabs:
        label(d, (x, 78), name, 14, BLUE if name == "Home" else TEXT, name == "Home")
    d.line((44, 105, 1236, 105), fill="#e0e5eb", width=1)
    commands = ["New Email", "Delete", "Archive", "Reply", "Forward", "Rules", "Move"]
    x = 65
    for cmd in commands:
        label(d, (x, 119), cmd, 13, TEXT, cmd == "Rules" and section in (3,4))
        x += 112 if cmd not in ("New Email", "Forward") else 120
    # left navigation pane
    d.rectangle((44, 152, 250, 640), fill="#f5f7fa")
    label(d, (62, 169), "FAVORITES", 12, MUTED, True)
    nav = [("Inbox", 202), ("Sent Items", 242), ("Drafts", 282), ("Deleted Items", 322), (PST_NAME, 390), ("   Inbox", 430), ("   Sent Items", 470)]
    for name, y in nav:
        if name == PST_NAME:
            label(d, (61, y), "▾", 15, BLUE, True)
            label(d, (82, y), name, 13, TEXT, True)
        else:
            label(d, (76, y), name, 13, TEXT)
    # message list
    d.rectangle((251, 152, 565, 640), fill=WHITE, outline="#e2e7ed")
    label(d, (270, 170), "Inbox", 18, TEXT, True)
    label(d, (270, 202), "Focused     Other", 12, MUTED)
    for i, (sender, subject) in enumerate([
        ("Operations Team", "Weekly status update"),
        ("IT Service Desk", "Maintenance notification"),
        ("Project Office", "Meeting notes"),
        ("Facilities", "Building access update"),
        ("Training Sample", "PST archive test message"),
    ]):
        y = 235 + i*72
        if i == 0:
            d.rectangle((252, y-7, 564, y+62), fill="#edf5fc")
        label(d, (270, y), sender, 12, TEXT, True)
        label(d, (270, y+23), subject[:34], 12, MUTED)
        d.line((263, y+59, 555, y+59), fill="#edf0f3")
    # reading pane
    d.rectangle((566, 152, 1236, 640), fill=WHITE)
    label(d, (594, 176), "Weekly status update", 20, TEXT, True)
    label(d, (594, 212), "From: Operations Team", 12, MUTED)
    label(d, (594, 236), "To: Alex Morgan", 12, MUTED)
    d.line((592, 263, 1205, 263), fill="#e5e9ef")
    label(d, (594, 290), "Hello,", 14)
    label(d, (594, 322), "Please review the latest project status", 14)
    label(d, (594, 348), "and confirm the next maintenance window.", 14)
    label(d, (594, 398), "Regards,", 14)
    label(d, (594, 425), "Operations Team", 14)
    # status bar and Windows taskbar
    d.rectangle((44, 641, 1236, 665), fill="#f6f7f9")
    label(d, (60, 646), "All folders are up to date", 11, MUTED)
    d.rectangle((0, 687, W, H), fill="#202b3a")
    label(d, (25, 693), "⊞", 18, WHITE, True)
    label(d, (66, 695), "Search", 12, "#dbe3ed")
    label(d, (1125, 695), "ENG   10:24 AM", 11, "#dbe3ed")
    # top-left section indicator
    rect(d, (890, 80, 1208, 105), "#edf4fb", None, 5)
    label(d, (903, 85), section, 12, BLUE, True)

def draw_account_settings(d, stage):
    # Modal window
    rect(d, (300, 165, 1000, 585), "#fbfcfd", "#8292a3", 8, 2)
    d.rectangle((302,167,998,207), fill="#f0f3f6")
    label(d, (322, 178), "Account Settings", 16, TEXT, True)
    label(d, (945, 178), "×", 18, MUTED)
    for i, tab in enumerate(["Email", "Data Files", "RSS Feeds", "SharePoint Lists"]):
        x = 330 + i*130
        if (stage == "data_files" or stage in ("add_dialog", "new_pst", "data_files_added", "default_confirm", "default_done", "data_files_default")) and tab == "Data Files":
            d.rectangle((x-8, 215, x+112, 248), fill=WHITE, outline="#aab6c2")
            label(d, (x, 224), tab, 13, BLUE, True)
        else:
            label(d, (x, 224), tab, 13, MUTED)
    d.line((322, 249, 975, 249), fill=BORDER)
    label(d, (330, 270), "Name", 12, MUTED, True)
    label(d, (620, 270), "Location", 12, MUTED, True)
    rowfill = LIGHT_BLUE if stage in ("data_files_added", "default_confirm", "default_done", "data_files_default") else WHITE
    d.rectangle((326, 295, 970, 345), fill=rowfill, outline="#b9c5d1")
    label(d, (342, 311), "Outlook Data File", 13)
    label(d, (620, 311), "Outlook\\outlook.ost", 13)
    d.rectangle((326, 352, 970, 402), fill="#f7fafc", outline="#d6dde5")
    label(d, (342, 368), PST_NAME, 13, TEXT, True)
    label(d, (620, 368), "Documents\\Outlook Files\\" + PST_NAME, 12, MUTED)
    if stage in ("default_confirm", "default_done", "data_files_default"):
        label(d, (342, 412), "Default data file", 12, GREEN, True)
    button(d, (760, 505, 865, 543), "Add...")
    button(d, (875, 505, 970, 543), "Set as Default", primary=True)
    button(d, (875, 548, 970, 578), "Close")


def draw_rules_dialog(d, stage):
    rect(d, (270, 125, 1030, 615), "#fbfcfd", "#7d8da0", 8, 2)
    d.rectangle((272,127,1028,168), fill="#f0f3f6")
    label(d, (292, 139), "Rules and Alerts", 16, TEXT, True)
    label(d, (980, 138), "×", 18, MUTED)
    label(d, (300, 184), "E-mail Rules", 14, TEXT, True)
    label(d, (300, 214), "Apply changes to this folder:", 12, MUTED)
    rect(d, (490, 205, 820, 235), WHITE, BORDER, 3)
    label(d, (503, 211), "Inbox - alex.morgan@example.com", 12)
    d.line((292, 250, 1005, 250), fill=BORDER)
    label(d, (302, 270), "Rule (applied in the order shown)", 12, MUTED, True)
    for idx, name in enumerate(["Inbox Archive Rule", "Sent Items Archive Rule"]):
        y = 310 + idx*42
        checkbox(d, 310, y, name, checked=(stage == "rules_list" or idx == 0))
    button(d, (300, 535, 390, 570), "New Rule", primary=True)
    button(d, (890, 570, 1000, 600), "Apply", primary=True)
    button(d, (780, 570, 880, 600), "Cancel")


def draw_wizard(d, stage, typed_progress=1.0):
    rect(d, (300, 105, 1000, 620), "#fbfcfd", "#7d8da0", 8, 2)
    d.rectangle((302,107,998,150), fill="#f0f3f6")
    label(d, (322, 120), "Rules Wizard", 16, TEXT, True)
    label(d, (340, 175), "What do you want to do with the message?", 17, TEXT, True)
    if stage in ("wizard_start", "wizard_received", "wizard_sent", "wizard_conditions"):
        label(d, (340, 218), "Step 1: Select a template or start from a blank rule", 12, MUTED)
        checkbox(d, 350, 270, "Apply rule on messages I receive", checked=stage in ("wizard_received", "wizard_conditions"))
        checkbox(d, 350, 310, "Apply rule on messages I send", checked=stage == "wizard_sent")
        checkbox(d, 350, 365, "from people or public group", checked=False)
        checkbox(d, 350, 405, "with specific words in the subject", checked=False)
        if stage == "wizard_conditions":
            rect(d, (330, 455, 950, 500), "#fff8e5", "#e2c46e", 5)
            label(d, (345, 468), "No conditions selected: this rule applies to all matching messages.", 12, "#795b12")
    elif stage in ("wizard_action", "wizard_sent_action"):
        label(d, (340, 218), "Step 2: Select the action(s)", 12, MUTED)
        checkbox(d, 350, 270, "move it to the specified folder", checked=stage == "wizard_action")
        checkbox(d, 350, 310, "move a copy to the specified folder", checked=stage == "wizard_sent_action")
        checkbox(d, 350, 365, "assign it to the category category", checked=False)
        checkbox(d, 350, 405, "flag the message for follow up", checked=False)
        label(d, (350, 470), "Edit the rule description below to choose a destination folder.", 12, MUTED)
    else:
        label(d, (340, 218), "Step 3: Choose the destination folder", 12, MUTED)
        label(d, (360, 270), "▾  Archive_1126.pst", 14, TEXT, True)
        folder = "Inbox" if stage == "folder_picker_inbox" else "Sent Items"
        rect(d, (390, 305, 850, 355), LIGHT_BLUE, BLUE, 4, 1)
        label(d, (420, 320), "▸  " + folder, 14, BLUE, True)
        label(d, (360, 385), "Rule name:", 13, MUTED)
        rule_name = "Inbox Archive Rule" if folder == "Inbox" else "Sent Items Archive Rule"
        rect(d, (470, 378, 850, 414), WHITE, BORDER, 3)
        label(d, (485, 387), rule_name, 13)
    button(d, (720, 565, 800, 595), "< Back")
    button(d, (810, 565, 890, 595), "Next >", primary=True)
    button(d, (900, 565, 980, 595), "Finish", primary=stage.startswith("rule_saved"))


def draw_file_menu(d):
    rect(d, (44, 65, 355, 410), WHITE, "#aab5c1", 5, 1)
    menu = [("Info", 100), ("Open & Export", 145), ("Save As", 190), ("Account Settings", 235), ("Options", 280), ("Exit", 325)]
    for text, y in menu:
        if text == "Account Settings":
            d.rectangle((48, y-7, 350, y+27), fill=LIGHT_BLUE)
        label(d, (72, y), text, 14, TEXT, text == "Account Settings")

def draw_rules_menu(d):
    rect(d, (315, 140, 560, 285), WHITE, "#aab5c1", 5, 1)
    label(d, (338, 160), "Manage Rules and Alerts...", 13, TEXT, True)
    label(d, (338, 205), "Create Rule...", 13)
    label(d, (338, 250), "Manage Quick Steps", 13)

def draw_new_pst_dialog(d, stage, progress):
    rect(d, (380, 245, 900, 485), "#fbfcfd", "#7d8da0", 8, 2)
    d.rectangle((382,247,898,285), fill="#f0f3f6")
    label(d, (402, 258), "Create or Open Outlook Data File", 15, TEXT, True)
    label(d, (410, 315), "File name:", 13, TEXT)
    rect(d, (500, 305, 860, 345), WHITE, BLUE if stage == "new_pst" else BORDER, 4, 1)
    shown = PST_NAME[:max(0, min(len(PST_NAME), int(len(PST_NAME)*progress)))]
    label(d, (512, 316), shown, 15)
    label(d, (410, 365), "Save as type: Outlook Data File (*.pst)", 12, MUTED)
    button(d, (690, 425, 765, 460), "OK", primary=True)
    button(d, (780, 425, 865, 460), "Cancel")

def draw_default_confirm(d):
    rect(d, (420, 275, 865, 455), WHITE, "#7d8da0", 7, 2)
    label(d, (450, 305), "Default Data File", 17, TEXT, True)
    label(d, (450, 350), "The selected data file will be set as default.", 13)
    label(d, (450, 375), "Continue?", 13)
    button(d, (660, 405, 745, 438), "Yes", primary=True)
    button(d, (755, 405, 840, 438), "No")

def draw_folder_picker(d, stage):
    rect(d, (385, 190, 895, 545), WHITE, "#7d8da0", 7, 2)
    label(d, (410, 212), "Select Folder", 16, TEXT, True)
    label(d, (425, 265), "▾  Archive_1126.pst", 14, TEXT, True)
    folder = "Inbox" if stage == "folder_picker_inbox" else "Sent Items"
    label(d, (465, 310), "▸  Deleted Items", 13, MUTED)
    rect(d, (450, 345, 840, 390), LIGHT_BLUE, BLUE, 4, 1)
    label(d, (480, 358), "▸  " + folder, 14, BLUE, True)
    button(d, (700, 490, 785, 525), "OK", primary=True)
    button(d, (795, 490, 880, 525), "Cancel")

def draw_testing(d, stage):
    # Overlay a realistic test message or result toast.
    if stage == "test_incoming":
        rect(d, (650, 280, 1150, 500), WHITE, "#93a3b4", 8, 2)
        label(d, (680, 305), "New message received", 17, TEXT, True)
        label(d, (680, 350), "From: training@example.com", 13, MUTED)
        label(d, (680, 380), "Subject: PST archive test message", 14, TEXT)
        label(d, (680, 425), "Inbox Archive Rule is processing this message...", 12, BLUE)
    elif stage == "test_inbox_done":
        rect(d, (260, 365, 880, 510), "#f4fbf6", "#7eb18d", 8, 2)
        label(d, (290, 390), "✓  Inbox archive verified", 19, GREEN, True)
        label(d, (290, 440), "Archive_1126.pst  >  Inbox", 15, TEXT)
        label(d, (290, 475), "PST archive test message", 13, MUTED)
    elif stage == "test_sent":
        rect(d, (610, 250, 1150, 535), WHITE, "#93a3b4", 8, 2)
        label(d, (640, 275), "New Message", 18, TEXT, True)
        label(d, (640, 320), "To: training@example.com", 13)
        label(d, (640, 360), "Subject: Sent items archive test", 13)
        d.line((640, 390, 1115, 390), fill=BORDER)
        label(d, (640, 415), "This is a sample message for the archive test.", 13)
        button(d, (1020, 475, 1115, 510), "Send", primary=True)
    elif stage == "test_sent_done":
        rect(d, (260, 365, 900, 520), "#f4fbf6", "#7eb18d", 8, 2)
        label(d, (290, 390), "✓  Sent Items archive verified", 19, GREEN, True)
        label(d, (290, 440), "Archive_1126.pst  >  Sent Items", 15, TEXT)
        label(d, (290, 475), "Sent items archive test", 13, MUTED)

def draw_final(d):
    rect(d, (320, 225, 960, 530), WHITE, "#a9b7c5", 12, 2)
    label(d, (390, 270), "CONFIGURATION COMPLETED", 27, GREEN, True)
    label(d, (390, 330), "Archive_1126.pst", 21, NAVY, True)
    label(d, (390, 375), "Inbox  ✓       Sent Items  ✓", 18, TEXT)
    label(d, (390, 445), "Thank you for watching", 22, BLUE, True)

def draw_cursor(d, x, y, click=False):
    # Pointer with dark outline and white fill.
    pts = [(x, y), (x+2, y+27), (x+9, y+20), (x+17, y+34), (x+23, y+31), (x+15, y+17), (x+26, y+16)]
    d.polygon(pts, fill=WHITE, outline="#111827")
    d.line(pts + [pts[0]], fill="#111827", width=2)
    if click:
        d.ellipse((x-12, y-12, x+34, y+34), outline=BLUE, width=3)

def step_at(scene, t):
    elapsed = 0.0
    for idx, step in enumerate(scene["steps"]):
        caption, state, target, duration = step
        if t < elapsed + duration or idx == len(scene["steps"])-1:
            return idx, step, max(0.0, t-elapsed), elapsed
        elapsed += duration
    return len(scene["steps"])-1, scene["steps"][-1], 0.0, elapsed

def smooth(a):
    a = max(0.0, min(1.0, a))
    return a*a*(3-2*a)

def render_frame(scene_index, scene, t):
    im = Image.new("RGB", (W, H), "#dce7f2")
    d = ImageDraw.Draw(im)
    draw_outlook_base(d, "OUTLOOK CLASSIC TRAINING")
    idx, step, local_t, elapsed = step_at(scene, t)
    caption, state, target, step_duration = step
    # Determine the previous cursor target, then animate toward this step's target.
    prev = scene["steps"][idx-1][2] if idx > 0 else (650, 360)
    travel = min(1.0, local_t / min(1.5, max(0.7, step_duration*0.28)))
    q = smooth(travel)
    cx = prev[0] + (target[0]-prev[0])*q
    cy = prev[1] + (target[1]-prev[1])*q
    clicked = local_t > min(1.5, max(0.7, step_duration*0.28)) and local_t < min(2.2, step_duration*0.42)

    if scene_index == 1:
        if state in ("file_menu", "account_menu"):
            draw_file_menu(d)
        if state in ("data_files", "add_dialog", "new_pst", "data_files_added", "default_confirm", "default_done", "data_files_default"):
            draw_account_settings(d, state)
        if state == "add_dialog":
            rect(d, (420, 270, 860, 450), WHITE, "#7d8da0", 7, 2)
            label(d, (450, 295), "Add Account", 16, TEXT, True)
            label(d, (450, 345), "Choose the type of data file to add.", 13)
            button(d, (670, 390, 835, 425), "Outlook Data File", primary=True)
        if state == "new_pst":
            draw_new_pst_dialog(d, state, local_t / max(1, step_duration-2))
        if state == "data_files_added":
            rect(d, (326, 352, 970, 402), LIGHT_BLUE, BLUE, 3, 2)
    elif scene_index == 2:
        draw_account_settings(d, state)
        if state == "default_confirm":
            draw_default_confirm(d)
    elif scene_index in (3, 4):
        if state in ("rules_menu",):
            draw_rules_menu(d)
        elif state in ("rules_alerts", "rules_list"):
            draw_rules_dialog(d, state)
        elif state in ("wizard_start", "wizard_received", "wizard_sent", "wizard_conditions", "wizard_action", "wizard_sent_action", "rule_saved_inbox", "rule_saved_sent"):
            draw_wizard(d, state)
            if state.startswith("rule_saved"):
                rect(d, (410, 280, 890, 420), "#f4fbf6", "#7eb18d", 7, 2)
                label(d, (445, 310), "Rule successfully created", 18, GREEN, True)
                label(d, (445, 355), "Inbox Archive Rule" if state.endswith("inbox") else "Sent Items Archive Rule", 15, TEXT)
        elif state in ("folder_picker_inbox", "folder_picker_sent"):
            draw_wizard(d, state)
            draw_folder_picker(d, state)
    elif scene_index == 5:
        draw_testing(d, state)
        if state in ("test_inbox_done", "test_sent_done"):
            # Highlight the matching PST subfolder in the navigation tree.
            y = 426 if state == "test_inbox_done" else 466
            rect(d, (52, y-8, 240, y+25), LIGHT_BLUE, BLUE, 4, 1)
    elif scene_index == 6:
        draw_final(d)

    # Header title and bottom instructional caption
    rect(d, (265, 156, 1218, 190), "#f7f9fc", None, 4)
    label(d, (282, 163), scene["title"], 17, NAVY, True)
    # Step caption bar, with a progress indicator
    d.rectangle((0, 620, W, 686), fill="#15243a")
    label(d, (28, 631), f"STEP {idx+1:02d}/{len(scene['steps']):02d}", 12, "#9ec7ef", True)
    label(d, (185, 629), caption, 18, WHITE, True)
    # progress bar for current scene
    scene_progress = min(1.0, t / scene["duration"])
    d.rectangle((0, 682, int(W*scene_progress), 686), fill=BLUE)
    draw_cursor(d, int(cx), int(cy), clicked)
    return im

def run(cmd, **kwargs):
    print("+", " ".join(map(str, cmd)), flush=True)
    subprocess.run(list(map(str, cmd)), check=True, **kwargs)

def make_audio(scene, index):
    wav = WORK / f"voice_{index:02d}.wav"
    text = scene["narration"]
    run(["espeak-ng", "-v", "en-us", "-s", "138", "-a", "165", "-w", wav, text])
    return wav

def make_video_segment(scene, index):
    duration = scene["duration"]
    target = WORK / f"scene_{index:02d}.mp4"
    cmd = [
        "ffmpeg", "-y", "-loglevel", "error",
        "-f", "rawvideo", "-pix_fmt", "rgb24", "-s", f"{W}x{H}",
        "-r", str(FPS), "-i", "-", "-an", "-c:v", "libx264",
        "-preset", "veryfast", "-crf", "23", "-pix_fmt", "yuv420p",
        "-movflags", "+faststart", str(target)
    ]
    proc = subprocess.Popen(cmd, stdin=subprocess.PIPE)
    total_frames = duration * FPS
    try:
        for frame_no in range(total_frames):
            t = frame_no / FPS
            frame = render_frame(index, scene, t)
            proc.stdin.write(frame.tobytes())
        proc.stdin.close()
        code = proc.wait()
        if code:
            raise subprocess.CalledProcessError(code, cmd)
    except Exception:
        if proc.stdin and not proc.stdin.closed:
            proc.stdin.close()
        proc.kill()
        raise
    return target

def main():
    for executable in ("ffmpeg", "ffprobe", "espeak-ng"):
        if not shutil.which(executable):
            raise SystemExit(f"Missing required executable: {executable}")
    total = sum(scene["duration"] for scene in SCENES)
    if total != TOTAL_SECONDS:
        raise SystemExit(f"Scene durations must total {TOTAL_SECONDS}, got {total}")

    videos, audios = [], []
    for i, scene in enumerate(SCENES):
        print(f"Rendering scene {i+1}/{len(SCENES)}: {scene['title']}", flush=True)
        videos.append(make_video_segment(scene, i))
        audios.append(make_audio(scene, i))

    concat_file = WORK / "concat.txt"
    concat_file.write_text("".join(f"file '{p.as_posix()}'\n" for p in videos), encoding="utf-8")
    silent_video = WORK / "video_only.mp4"
    run(["ffmpeg", "-y", "-loglevel", "error", "-f", "concat", "-safe", "0", "-i", concat_file, "-c", "copy", silent_video])

    # Concatenate voice clips, padding each one to its scene duration.
    audio_inputs = []
    filters = []
    for i, (scene, wav) in enumerate(zip(SCENES, audios)):
        audio_inputs += ["-i", str(wav)]
        filters.append(f"[{i}:a]apad,atrim=0:{scene['duration']},asetpts=PTS-STARTPTS[a{i}]")
    filters.append("".join(f"[a{i}]" for i in range(len(audios))) + f"concat=n={len(audios)}:v=0:a=1[outa]")
    narration = WORK / "narration.wav"
    run(["ffmpeg", "-y", "-loglevel", "error", *audio_inputs, "-filter_complex", ";".join(filters), "-map", "[outa]", "-ar", "44100", "-ac", "2", narration])

    final = OUT / "Outlook_Classic_Training.mp4"
    run(["ffmpeg", "-y", "-loglevel", "error", "-i", silent_video, "-i", narration,
         "-map", "0:v:0", "-map", "1:a:0", "-c:v", "copy", "-c:a", "aac", "-b:a", "128k",
         "-t", str(TOTAL_SECONDS), "-movflags", "+faststart", final])
    # Verify playable container and exact duration.
    probe = subprocess.check_output(["ffprobe", "-v", "error", "-show_entries", "format=duration,size", "-of", "default=noprint_wrappers=1", str(final)], text=True)
    print("\nVIDEO VERIFIED\n" + probe, flush=True)
    print(f"Created: {final}", flush=True)

if __name__ == "__main__":
    main()
