import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# --- Helper Functions ---

def ordinal(n):
    if 11 <= (n % 100) <= 13:
        suffix = "th"
    else:
        suffix = {1:"st", 2:"nd", 3:"rd"}.get(n % 10, "th")
    return str(n) + suffix

def format_date(val):
    """Convert to 'Dth Month'."""
    if pd.isna(val) or str(val).strip() == "":
        return ""
    if isinstance(val, datetime):
        dt = val
    else:
        for fmt in ("%Y-%m-%d %H:%M:%S", "%d %B %Y", "%d-%m-%Y %H:%M:%S"):
            try:
                dt = datetime.strptime(str(val).strip(), fmt)
                break
            except:
                continue
        else:
            return str(val)
    return f"{ordinal(dt.day)} {dt.strftime('%B')}"

def split_items(text):
    t = str(text).strip()
    if not t:
        return []
    if "\n" in t:
        parts = t.split("\n")
    else:
        parts = t.split(",")
    return [p.strip() for p in parts if p.strip()]

def count_items(text):
    return len(split_items(text))

def count_direct_items(text):
    return len([p for p in split_items(text) if "(relationtype = direct)" in p.lower()])

def preserve(val):
    s = str(val)
    return s if s == " " else s.strip()

def build_summary(location, online, ci, bc, nonbc):
    parts = [preserve(location), online.strip()]
    
    ci_count = count_items(ci)
    if ci_count > 0:
        parts.append("1 CI" if ci_count == 1 else f"{ci_count} CIs")
    
    bc_direct = count_direct_items(bc)
    if bc_direct > 0:
        parts.append(f"{bc_direct} BC (Direct)")
    
    nonbc_direct = count_direct_items(nonbc)
    if nonbc_direct > 0:
        parts.append(f"{nonbc_direct} NON BC (Direct)")
    
    return ", ".join([p for p in parts if p])

# --- Excel Generation ---

def generate_formatted_excel(df):
    output = BytesIO()
    df.fillna(" ", inplace=True)
    df.columns = df.columns.str.strip()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        ws = wb.add_worksheet("Output Final")

        bold = wb.add_format({'bold': True, 'text_wrap': True, 'bg_color': 'white', 'font_color': 'black', 'font_size': 12})
        norm = wb.add_format({'text_wrap': True, 'bg_color': 'white', 'font_color': 'black', 'font_size': 12})

        ws.set_column(0, 0, 50)
        ws.set_column(1, 1, 30)
        ws.set_column(2, 2, 50)

        row = 0
        for _, r in df.iterrows():
            # Column 1
            fs = format_date(r.PlannedStart)
            fe = format_date(r.PlannedEnd)
            if fs and fe:
                date_line = fs if fs == fe else f"{fs} – {fe} {r.PlannedStart.year}"
            else:
                date_line = ""
            title = preserve(r.Title)
            summary = build_summary(r.Location, r["OnLine/Outage"], r.CI, r.BC, r.NONBC)
            bg = preserve(r.BusinessGroups)

            c1 = [
                " ", bold, date_line,
                norm, "\n\n",
                norm, title,
                norm, "\n\n",
                norm, summary,
                norm, "\n\n",
                norm, bg
            ]
            ws.write_rich_string(row, 0, *c1)

            # Column 2
            cid = preserve(r.ChangeId)
            f4f = preserve(r.F4F)
            risk = preserve(r.RiskLevel.replace("SHELL_", "", 1).capitalize())
            if f4f == " " and not risk:
                c2 = cid
            else:
                first = cid if f4f == " " else (f"{cid}/{f4f}" if cid != " " else f4f)
                c2 = f"{first}\n{risk}"
            ws.write_rich_string(row, 1, " ", norm, c2)

            # Column 3
            bc_items = split_items(r.BC)
            trading = [
                p.replace("(RelationType = Direct)", "", 1).strip()
                for p in bc_items
                if "(relationtype = direct)" in p.lower() and p.upper().startswith("ST")
            ]
            other_bc_direct = [
                p.replace("(RelationType = Direct)", "", 1).strip()
                for p in bc_items
                if "(relationtype = direct)" in p.lower() and not p.upper().startswith("ST")
            ]
            nonbc_items = split_items(r.NONBC)
            nonbc_trading = [
                p.replace("(RelationType = Direct)", "", 1).strip()
                for p in nonbc_items
                if "(relationtype = direct)" in p.lower() and p.upper().startswith("ST")
            ]

            if trading:
                parts3 = [
                    " ", bold, "Trading assets in scope: ", norm, "Yes",
                    norm, "\n\n", bold, "Trading BC Apps: ", norm, ", ".join(trading),
                    norm, "\n\n", bold, "Other BC Apps: ", norm,
                    (", ".join(other_bc_direct) if other_bc_direct else "No")
                ]
            elif nonbc_trading:
                parts3 = [
                    " ", bold, "Trading assets in scope: ", norm, "Yes (NON BC)",
                    norm, "\n\n", bold, "Other BC Apps: ", norm, ", ".join(nonbc_trading)
                ]
            else:
                parts3 = [
                    " ", bold, "Trading assets in scope: ", norm, "No",
                    norm, "\n\n", bold, "Other BC Apps: ", norm,
                    (", ".join(other_bc_direct) if other_bc_direct else "No")
                ]

            ws.write_rich_string(row, 2, *parts3)

            row += 1

    output.seek(0)
    return output

# --- Streamlit UI ---

st.title("Change Formatter App")
uploaded = st.file_uploader("Upload your Changes Excel file", type=["xlsx","xls"])
if uploaded:
    df = pd.read_excel(uploaded)
    df.fillna(" ", inplace=True)
    df.columns = df.columns.str.strip()
    out = generate_formatted_excel(df)
    st.download_button("📥 Download Formatted Output", out,
                       file_name="output_final.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
