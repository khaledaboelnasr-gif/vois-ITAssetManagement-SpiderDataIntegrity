"""Generate synthetic VCAMS exports (previous + current month) and reference tabs. No real data.

Usage: python gen_synthetic.py [assets] [--minimal] > synthetic.json
Writes {tab name: rows (header first)} for:
  "2026-09 Previous Export", "2026-10 Current Export"   VCAMS layout (the 13 standard headers + optional columns)
  "PO Validation"                                        user-supplied PO rows
  "Users (Entra)", "Intune Devices"                      optional references
--minimal keeps only the 13 standard VCAMS headers (no user, hostname, location or company columns).

Planted, so every sheet has something to show (scaled with the asset count):
  new Bags (most on PO 1600513572, some with a blank PurchaseNo), assets removed from the current export, serial number
  changes (changed, blanked, filled in, and a case-only difference that is not a change), status changes
  (some to Scrapped / Lost, most excluded), a scrapped asset back in operation, status spelling variants that are NOT
  changes (" Reserved ", "SCRAPPED"), unknown statuses, a duplicate asset number, placeholder serials, leavers,
  devices missing from Intune, a scrapped notebook still checking in, old idle stock.
"""
import json, random, sys
from datetime import datetime, timedelta

random.seed(11)
args = [a for a in sys.argv[1:] if not a.startswith("--")]
N = int(args[0]) if args else 5000
MINIMAL = "--minimal" in sys.argv
PREV, CUR = datetime(2026, 9, 8), datetime(2026, 10, 8)

STD = ["Asset no.", "Serial no.", "Asset status", "Asset Type", "Model", "Purchase date", "PurchaseNo", "Last scan",
       "Scandate Age in Days", "Function Unit", "AssetCreationDate", "Creation Age in Days", "AssetUpdateDate"]
EXTRA = ["Hostname", "Assigned To", "User Email", "Location", "Company", "Comment", "Status comment"]
HEADER = STD + ([] if MINIMAL else EXTRA)
TYPES = [  # type, weight, models, needs user, scanned
    ("Notebook", 30, ["Latitude 5440", "EliteBook 840 G9", "ThinkPad T14"], True, True),
    ("Desktop", 4, ["OptiPlex 7010"], True, True), ("Monitor", 14, ["P2422H", "E24 G5"], False, False),
    ("Zubehör-Client", 4, ["WD19S Dock"], False, False), ("Headset", 10, ["Evolve2 40", "Blackwire 5220"], True, False),
    ("Bag", 17, ["Example Bag"], False, False), ("Mouse", 6, ["MS116"], False, False), ("Mobile", 5, ["iPhone 14", "Galaxy A54"], True, False),
    ("Keyboard", 6, ["KB216"], False, False), ("IP Phone", 1, ["CP-8845"], False, False), ("VC Device", 0.5, ["Room Kit"], False, False),
    ("Printer", 0.5, ["Xerox C405"], False, False), ("Virtual Machine", 0.3, ["VM-Std"], False, False),
    ("Docking Station", 0.2, ["TB4 Dock"], False, False)]  # last one is not in the standard list
STATUSES = [("In operation", 66), ("In stock", 12), ("Reserved", 2), ("Ready for use", 2), ("To install", 1), ("Building", 1),
            ("Ordered", 1), ("Disabled", 2), ("Authorised inactive", 2), ("Scrapped", 9), ("Lost", 0.5), ("Marketed", 1.5)]
MARKETS = [("Vodafone Egypt Shared Services", "EG", ["Cairo - Smart Village B1", "Cairo - Smart Village B2", "Alexandria"], 0.5),
           ("Vodafone India Services", "IN", ["Pune - Magarpatta", "Bangalore - Bagmane"], 0.4),
           ("Vodafone Romania Shared Services", "RO", ["Bucharest"], 0.1)]

def serial(d):
    return round((d - datetime(1899, 12, 30)).total_seconds() / 86400, 4) if d else ""

users = [(f"Test User {i:05d}", f"test.user{i:05d}@vodafone.example") for i in range(1, max(2600, N // 2) + 1)]
leavers = set(range(1, len(users) + 1, 37))

def make(i, when):
    typ, _, models, needs_user, scanned = random.choices(TYPES, weights=[t[1] for t in TYPES])[0]
    status = random.choices([s[0] for s in STATUSES], weights=[s[1] for s in STATUSES])[0]
    company, code, sites, _ = random.choices(MARKETS, weights=[m[3] for m in MARKETS])[0]
    created = when - timedelta(days=random.randint(20, 6 * 365))
    bought = created - timedelta(days=random.randint(1, 20))
    updated = created + timedelta(days=random.randint(0, max(1, (when - created).days - 1)), hours=random.randint(8, 17))
    u = random.choice(users) if needs_user and status == "In operation" else ("", "")
    a = {"Asset no.": f"AST-{i:07d}", "Serial no.": "" if typ in ("Bag", "Virtual Machine") else f"{code}{i:08d}", "Asset status": status,
         "Asset Type": typ, "Model": random.choice(models), "Purchase date": bought, "PurchaseNo": f"45{random.randint(10000000, 10099999)}",
         "Last scan": None, "Function Unit": f"FU-{code}-{random.randint(1, 12):02d}", "AssetCreationDate": created, "AssetUpdateDate": updated,
         "Hostname": f"{code}L{i:07d}" if typ in ("Notebook", "Desktop") else "", "Assigned To": u[0], "User Email": u[1],
         "Location": random.choice(sites), "Company": company, "Comment": "", "Status comment": "", "_scanned": scanned}
    if scanned and status == "In operation":
        a["Last scan"] = when - timedelta(days=random.choice([1, 2, 3, 5, 8, 12, 20, 70]), hours=random.randint(0, 9))
    if i % 97 == 0 and a["Serial no."]:
        a["Serial no."] = random.choice(["N/A", "To be filled by O.E.M.", "0"])
    return a

def row(a, when):
    out = []
    for h in HEADER:
        if h == "Scandate Age in Days":
            out.append((when - a["Last scan"]).days if a["Last scan"] else "")
        elif h == "Creation Age in Days":
            out.append((when - a["AssetCreationDate"]).days)
        elif h in ("Purchase date", "Last scan", "AssetCreationDate", "AssetUpdateDate"):
            out.append(serial(a[h]))
        else:
            out.append(a[h])
    return out

prev = [make(i, PREV) for i in range(1, N + 1)]
cur = [dict(a) for a in prev]
# status spelling variants in the previous export: the same status, not a change
for a in prev[::211]:
    if a["Asset status"] == "Reserved": a["Asset status"] = " Reserved "
    if a["Asset status"] == "Scrapped": a["Asset status"] = "SCRAPPED"
k = max(1, N // 1000)
# status changes this month
for a in random.sample(cur, k=N // 25):
    a["AssetUpdateDate"] = CUR - timedelta(days=random.randint(1, 25), hours=random.randint(0, 8))
    s = a["Asset status"]
    if s in ("In operation", "Disabled", "In stock") and random.random() < 0.18:
        a["Asset status"] = "Lost" if random.random() < 0.03 else "Stolen" if random.random() < 0.02 else "Scrapped"
        a["Status comment"] = f"Write Off Memo as per INC{random.randint(10**11, 10**12 - 1)}"
        if random.random() < 0.3: a["Comment"] = "Pending investigation"
    elif s == "In operation" and random.random() < 0.02:
        a["Asset status"] = "Inactive"
    elif s in ("In stock", "Ready for use", "Reserved", "To install"):
        a["Asset status"] = "In operation"
        if a["Asset Type"] in ("Notebook", "Desktop", "Headset", "Mobile"):
            u = random.choice(users); a["Assigned To"], a["User Email"] = u
    elif s == "In operation":
        a["Asset status"] = random.choice(["In stock", "Disabled", "Authorised inactive"])
    elif s == "Ordered":
        a["Asset status"] = "In stock"
scr = [a for a in cur if a["Asset status"] == "Scrapped" and a["Asset Type"] == "Notebook"]
for a in scr[:2]:  # back from the dead
    a["Asset status"] = "In operation"
for a in scr[2:2 + 3 * k]:  # scrapped notebook still checking in after the change
    a["Last scan"] = CUR - timedelta(days=2)
cur[5]["Asset status"] = "Awaiting Triage"; cur[6]["Asset status"] = "in operations"
# serial number changes: corrected, filled in, blanked; plus a case/spacing-only difference that is NOT a change
withsn = [a for a in cur if a["Serial no."] and a["Serial no."] not in ("N/A", "To be filled by O.E.M.", "0")]
for a in withsn[100:100 + 6 * k]: a["Serial no."] = a["Serial no."][:-3] + "XYZ"
for a in withsn[200:200 + 2 * k]: a["Serial no."] = ""
for a in withsn[300:300 + 2 * k]: a["Serial no."] = a["Serial no."].lower() + " "
for a in [x for x in cur if x["Serial no."] in ("N/A", "0")][:k]: a["Serial no."] = f"FIX{random.randint(10000, 99999)}"
removed = set(a["Asset no."] for a in random.sample(cur, k=5 * k))
cur = [a for a in cur if a["Asset no."] not in removed]
# new bags: most on one PO, some with a blank PurchaseNo
for j in range(80 * k):
    a = make(900000 + j, CUR)
    a.update({"Asset no.": f"AST-9{j:06d}", "Asset Type": "Bag", "Model": "Example Bag", "Serial no.": "", "Asset status": "In stock",
              "Purchase date": CUR - timedelta(days=25), "AssetCreationDate": CUR - timedelta(days=22, hours=-10),
              "AssetUpdateDate": CUR - timedelta(days=21), "PurchaseNo": "1600513572" if j % 8 else "", "Last scan": None,
              "Hostname": "", "Assigned To": "", "User Email": "", "Location": "Cairo - Smart Village B1", "Company": MARKETS[0][0]})
    cur.append(a)
dup = dict(cur[10]); cur.append(dup)  # an asset number twice in the current export
# old idle notebook stock in Alexandria
for j in range(20):
    a = make(800000 + j, CUR)
    a.update({"Asset no.": f"AST-8{j:06d}", "Asset Type": "Notebook", "Model": "EliteBook 840 G5", "Asset status": "In stock",
              "Purchase date": CUR - timedelta(days=6 * 365), "AssetCreationDate": CUR - timedelta(days=6 * 365),
              "AssetUpdateDate": CUR - timedelta(days=420), "Assigned To": "", "User Email": "", "Location": "Alexandria",
              "Company": MARKETS[0][0], "Last scan": None})
    prev.append(dict(a)); cur.append(a)

data = {"2026-09 Previous Export": [HEADER] + [row(a, PREV) for a in prev],
        "2026-10 Current Export": [HEADER] + [row(a, CUR) for a in cur],
        "PO Validation": [["Purchase Date", "Purchase Number", "Model", "Quantity", "Raised By", "Physical Security Validation", "Spider Validation"],
                          [serial(CUR - timedelta(days=25)), "1600513572", "Bags", 70 * k, "USER-0001", "Yes", "Yes"],
                          [serial(CUR - timedelta(days=25)), "1600513572", "Mouses", 20 * k, "USER-0001", "Yes", "Yes"],
                          [serial(CUR - timedelta(days=20)), "1600599999", "Keyboards", 40, "USER-0002", "Pending", "Yes"]]}
if not MINIMAL:
    data["Users (Entra)"] = [["Display Name", "Email", "Account Enabled", "Leaving Date"]] + [
        [n, m, "False" if i in leavers else "True", (CUR - timedelta(days=30 + i % 150)).strftime("%Y-%m-%d") if i in leavers else ""]
        for i, (n, m) in enumerate(users, 1)]
    intune = [["Device name", "Serial number", "Last check-in", "Primary user UPN"]]
    for i, a in enumerate(cur):
        if a["Asset Type"] in ("Notebook", "Desktop") and a["Last scan"] and i % 23:
            intune.append([a["Hostname"], a["Serial no."], a["Last scan"].strftime("%Y-%m-%d %H:%M"), a["User Email"] if i % 41 else "someone.else@vodafone.example"])
    for j in range(30):
        intune.append([f"EGL9{j:06d}", f"EGZ99{j:05d}", (CUR - timedelta(days=j)).strftime("%Y-%m-%d %H:%M"), ""])
    data["Intune Devices"] = intune
json.dump(data, sys.stdout, ensure_ascii=False)
