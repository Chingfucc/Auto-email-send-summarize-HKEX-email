# Author: Donald Fu
# Date: 19 Nov 2025
###############################################################################
import win32com.client
import datetime
import re
import ctypes

outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")
inbox = mapi.GetDefaultFolder(6)

# === PRODUCTION: Use yesterday ===
# =============================================================================
# today = datetime.date.today()
# yesterday = today - datetime.timedelta(days=1)
# start = yesterday.strftime('%m/%d/%Y') + ' 12:00 AM'
# end   = today.strftime('%m/%d/%Y')   + ' 12:00 AM'
# report_date = yesterday.strftime('%d %b %Y')
# =============================================================================

# === TEST TODAY? Uncomment below ===
# start = today.strftime('%m/%d/%Y') + ' 12:00 AM'
# end   = (today + datetime.timedelta(days=1)).strftime('%m/%d/%Y') + ' 12:00 AM'
# report_date = today.strftime('%d %b %Y')

today = datetime.date.today()
today_start = today.strftime('%m/%d/%Y') + ' 12:00 AM'
tomorrow_start = (today + datetime.timedelta(days=1)).strftime('%m/%d/%Y') + ' 12:00 AM'

filter_str = (
    f"[ReceivedTime] >= '{today_start}' AND "
    f"[ReceivedTime] < '{tomorrow_start}' AND "
    f"[Subject] = 'HKEX News Alert' AND"
    f"[SenderEmailAddress] = 'donaldfu@phillip.com.hk'"
)

items = inbox.Items.Restrict(filter_str)
items.Sort("[ReceivedTime]", False)

# === CHECK IF ANY EMAIL FOUND ===
if items.Count == 0:
    ctypes.windll.user32.MessageBoxW(0, "No HKEX News Alert email found today.", "No Email", 1)
    raise SystemExit("No HKEX News Alert email today.")

messages = inbox.Items.Restrict(filter_str)
messages.Sort("[ReceivedTime]", False)

# Storage
grouped_content = {}
announcement_full_line = ""

for message in messages:
    body = message.Body

    # === STEP 1: CAPTURE FULL ANNOUNCEMENT LINE (EXACTLY AS HKEX WRITES IT) ===
    if "Announcement -" in body:
        start_pos = body.find("Announcement -")
        end_pos = body.find("Participant Circulars -", start_pos)
        if end_pos == -1:
            end_pos = body.find("You are receiving this alert", start_pos)
        if end_pos == -1:
            end_pos = len(body)

        raw_anno = body[start_pos:end_pos]

        # Fix only the broken parts: remove \ and join lines, but keep original wording
        raw_anno = raw_anno.replace("\\", "")
        raw_anno = re.sub(r'\s*\r?\n\s*', '', raw_anno)   # Join broken lines
        raw_anno = re.sub(r'\s{2,}', '', raw_anno).strip()  # Clean spaces

        # Find the URL and make it clickable, keep everything else exactly as-is
        def make_clickable(match):
            url = match.group(0)
            return f'<a href="{url}" style="color:blue; text-decoration:underline;">{url}</a>'
        
        announcement_full_line = re.sub(r'http[s]?://[^\s]+', make_clickable, raw_anno)

    # === STEP 2: PARTICIPANT CIRCULARS (your original working logic) ===
    circ_start = body.find("Participant Circulars -")
    if circ_start == -1:
        continue
    circ_end = body.find("You are receiving this alert")
    if circ_end == -1:
        circ_end = len(body)

    content_block = body[circ_start:circ_end]

    cat_match = re.search(r"Participant Circulars - (.+?)\s*\(", content_block)
    category = cat_match.group(1).strip() if cat_match else "Unknown"

# === BUILD FINAL EMAIL ===
html_body = f"""<html><body>
<font face="Arial" size="3"><b>HKEx Circulars – {today}</b></font><br><br>
<pre style="font-family:Arial; font-size:10pt; line-height:1.5;">
"""

# === 1. ANNOUNCEMENT – ONE PERFECT LINE (EXACTLY WHAT YOU WANTED) ===
if announcement_full_line:
    html_body += "<b>=== CORPORATE / REGULATORY ANNOUNCEMENT ===</b>\n\n"
    html_body += announcement_full_line + "\n\n\n"

# === 2. PARTICIPANT CIRCULARS ===
for category, contents in grouped_content.items():
    html_body += f"<b>--- {category} ---</b>\n\n"
    for content in contents:
        lines = content.splitlines()
        fixed_lines = []
        url_buffer = ""
        for line in lines:
            line = line.rstrip()
            
            # Collect URL fragments across broken lines
            if 'http' in line or url_buffer:
                url_buffer += line.replace("\\", "").replace(" ", "")
                if '.pdf' in url_buffer.lower():
                    # Full URL reconstructed → clean & make clickable
                    full_url = re.sub(r'\s+', '', url_buffer).split('?')[0]
                    pdf_name = full_url.split('/')[-1]
                    clickable_link = f'<a href="{full_url}" style="color:blue; text-decoration:underline;">{full_url}</a>'
                    fixed_lines.append(clickable_link)
                    url_buffer = ""
                continue

            # Reset if incomplete URL
            if url_buffer:
                url_buffer = ""

            fixed_lines.append(line)

        html_body += "\n".join(fixed_lines) + "\n\n"

html_body += """
</pre>
<hr>
</body></html>
"""

# === SEND EMAIL ===
mail = outlook.CreateItem(0)
mail.Subject = "HKEx Circulars"
mail.HTMLBody = html_body

# Set sender
for acc in outlook.Session.Accounts:
    if acc.SmtpAddress.lower() in ["audit@phillip.com.hk", "donaldfu@phillip.com.hk"]:
        mail.SendUsingAccount = acc
        break

# Recipients
mail.To = "donaldfu@phillip.com.hk"  # Test
# UNCOMMENT FOR FULL SEND:
# mail.To = "audit@phillip.com.hk;account@phillip.com.hk;assetmgt@phillip.com.hk;cs@phillip.com.hk;credit@phillip.com.hk;dealing@phillip.com.hk;futures@phillip.com.hk;ia@phillip.com.hk;it@phillip.com.hk;ipo@phillip.com.hk;option@phillip.com.hk;settlement@phillip.com.hk;ut@phillip.com.hk;bonds@phillip.com.hk"

# mail.Display(True)  # ← Use this to preview
mail.Send()

ctypes.windll.user32.MessageBoxW(0, "HKEx email sent – Announcement shown in one perfect line!", "Success", 1)