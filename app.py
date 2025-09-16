
import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
import os, re, base64, json, requests
import msal

st.set_page_config(page_title="Wareneingangsbestellung Rotogal", page_icon="📄", layout="wide")

st.title("📄 Wareneingangsbestellung Rotogal (Microsoft Cloud)")
st.caption("One‑page A4 PDF. English document. Upload to OneDrive and send email via Microsoft Graph.")

# ---------------- Helpers ----------------
def eur_fmt(x: float) -> str:
    try:
        v = float(x)
    except Exception:
        return ""
    s = f"{v:,.2f}"
    s = s.replace(",", "_").replace(".", ",").replace("_", ".")
    return s

def scale_mm(widths_mm, target_total_mm):
    s = sum(widths_mm)
    if s <= 0:
        return widths_mm
    f = target_total_mm / s
    return [w * f for w in widths_mm]

def clean(val):
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    sval = str(val).strip()
    if sval.lower() in ("nan", "none", "null"):
        return ""
    return sval

def sanitize_filename(name: str) -> str:
    name = (name or "").strip() or "supplier_order"
    safe = re.sub(r"[^A-Za-z0-9\-\_\s\.]", "", name)
    safe = re.sub(r"\s+", "_", safe)
    if not safe.lower().endswith(".pdf"):
        safe += ".pdf"
    return safe

# ---------------- Microsoft Graph helpers ----------------
def get_graph_token(tenant_id: str, client_id: str, client_secret: str):
    try:
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app = msal.ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)
        scope = ["https://graph.microsoft.com/.default"]
        result = app.acquire_token_silent(scope, account=None)
        if not result:
            result = app.acquire_token_for_client(scopes=scope)
        if "access_token" in result:
            return result["access_token"], None
        else:
            return None, result.get("error_description") or str(result)
    except Exception as e:
        return None, str(e)

def onedrive_upload_file(access_token: str, user_upn: str, folder_path: str, filename: str, file_bytes: bytes):
    folder_path = (folder_path or "").strip("/")
    if folder_path:
        url = f"https://graph.microsoft.com/v1.0/users/{user_upn}/drive/root:/{folder_path}/{filename}:/content"
    else:
        url = f"https://graph.microsoft.com/v1.0/users/{user_upn}/drive/root:/{filename}:/content"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/pdf"}
    r = requests.put(url, headers=headers, data=file_bytes, timeout=60)
    return r.status_code, r.text

def graph_send_mail(access_token: str, sender_upn: str, to_emails: list, subject: str, body_html: str, attachment_bytes: bytes = None, attachment_name: str = "order.pdf"):
    url = f"https://graph.microsoft.com/v1.0/users/{sender_upn}/sendMail"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    message = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body_html},
            "toRecipients": [{"emailAddress": {"address": e}} for e in to_emails],
        },
        "saveToSentItems": True
    }
    if attachment_bytes:
        content_b64 = base64.b64encode(attachment_bytes).decode("utf-8")
        message["message"]["attachments"] = [{
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": attachment_name,
            "contentType": "application/pdf",
            "contentBytes": content_b64
        }]
    r = requests.post(url, headers=headers, data=json.dumps(message), timeout=60)
    return r.status_code, r.text

# ---------------- Sidebar: meta ----------------
with st.sidebar:
    st.header("Order meta")
    company = st.text_input("Company", value="Rotogal GmbH")
    contact_person = st.text_input("Contact person", value="Maurice Vennegerts")
    tel = st.text_input("Phone", value="015221870004")
    email = st.text_input("E-mail", value="vennegerts@rotogal.de")
    order_no = st.text_input("Our Order No.", value="")
    your_order_ref = st.text_input("Your order ref. (internal)", value="")
    date_val = st.date_input("Date", value=datetime.today())

    st.markdown("---")
    st.subheader("Addresses")
    ship_to = st.text_area("Shipping address", value="")
    bill_to = st.text_area("Billing address", value="Rotogal GmbH\nDorfstr. 77\n49848 Wilsum\nGermany")

    st.markdown("---")
    st.subheader("VAT")
    vat_choice = st.selectbox("VAT ID", ["DE294750940 (Germany)", "ESN0300033H (Spain)"], index=0)
    if "ESN0300033H" in vat_choice:
        vat_id = "ESN0300033H"
        es_vat_rate = st.number_input("Spanish VAT rate (%)", min_value=0.0, max_value=30.0, value=21.0, step=0.5)
        vat_rate = es_vat_rate / 100.0
    else:
        vat_id = "DE294750940"
        vat_rate = 0.0

    st.markdown("---")
    st.subheader("Microsoft 365 (Graph)")
    tenant_id = st.text_input("Tenant ID", value=st.secrets.get("TENANT_ID", ""))
    client_id = st.text_input("Client ID", value=st.secrets.get("CLIENT_ID", ""))
    client_secret = st.text_input("Client Secret", value=st.secrets.get("CLIENT_SECRET", ""), type="password")
    graph_user_upn = st.text_input("User UPN for OneDrive/Email", value=st.secrets.get("GRAPH_USER_UPN", email), help="e.g., name@yourdomain.com")
    onedrive_folder = st.text_input("OneDrive target folder", value=st.secrets.get("ONEDRIVE_FOLDER", "Bestellungen/Rotogal"))
    email_to = st.text_input("Send email to", value=st.secrets.get("EMAIL_TO", ""), help="Comma‑separated list")
    auto_upload = st.checkbox("Auto‑upload PDF to OneDrive after export", value=False)
    auto_email = st.checkbox("Auto‑send email with PDF after export", value=False)

    st.markdown("---")
    st.subheader("Footer")
    footer_left = st.text_area("Footer (left)", value=(
        "Rotogal GmbH\nDorfstr. 77\nD-49848 Wilsum\nPhone " + tel + "\nFax\n\n"
        "Bank account:\nVolksbank Niedergrafschaft e.G.\nBIC: GENODEF1HOO\nIBAN: DE05280699262430498000\n\n"
        "Managing Director:\nGilbert Mommertz"
    ))
    footer_right_extra = st.text_area("Footer (right, extra lines)", value=(
        "Tax-No: 55/208/12604\nCommercial register: HRB 208659"
    ))

# ---------------- Presets ----------------
COLORS = ["Natural/White", "Blue", "Red", "Green", "Yellow", "Gray", "Black", "Orange", "Other (free text)"]
DRAIN_PLUG = ["None", "1\" drain", "1½\" drain", "2\" drain", "Other (free text)"]
WALL_BUILD = ["EPE", "PUR"]
PRODUCT_TYPES = {
    "Bins": ["Model", "Color", "Wall build", "Drain"],
    "Lids": ["Model", "Color", "Wall build"],
    "Buggies": ["Model", "Color"],
    "Pallets": ["Model", "Color"]
}

# ---------------- Session ----------------
if "order_lines" not in st.session_state:
    st.session_state.order_lines = []

def add_line(ptype, data):
    st.session_state.order_lines.append({"Product group": ptype, **data})

def reset_lines():
    st.session_state.order_lines = []

st.divider()

# ---------------- Line item builder ----------------
st.subheader("Add position")

colA, colB, colC = st.columns([1, 1, 1.2])
with colA:
    product_type = st.selectbox("Product group", list(PRODUCT_TYPES.keys()))
with colB:
    qty = st.number_input("Quantity", min_value=1, step=1, value=1)
with colC:
    net_price = st.number_input("Net price per item (EUR)", min_value=0.0, step=1.0, value=0.0, format="%.2f")

fields = PRODUCT_TYPES[product_type]

with st.form("line_form", clear_on_submit=True):
    cols = st.columns(3)
    values = {}
    for i, field in enumerate(fields):
        with cols[i % 3]:
            if field == "Model":
                values[field] = st.text_input(field, placeholder="e.g., BI-565, Hygiene pallet 1200x800 …")
            elif field == "Color":
                values[field] = st.selectbox(field, COLORS, index=1)
            elif field == "Wall build":
                values[field] = st.selectbox(field, WALL_BUILD, index=0)
            elif field == "Drain":
                values[field] = st.selectbox(field, DRAIN_PLUG, index=0)
            else:
                values[field] = st.text_input(field)

    remark = st.text_input("Note (optional)")
    submitted = st.form_submit_button("➕ Add position")
    if submitted:
        line_total = float(net_price) * int(qty)
        line = {
            "Quantity": int(qty),
            **values,
            "Note": remark,
            "Net price": float(net_price),
            "Total": line_total,
        }
        add_line(product_type, line)
        st.success("Position added.")

st.divider()

# ---------------- Order table + export ----------------
st.subheader("Order overview")
if len(st.session_state.order_lines) == 0:
    st.info("No positions added yet.")
else:
    df = pd.DataFrame(st.session_state.order_lines)
    df_display = df.copy()
    if "Net price" in df_display:
        df_display["Net price"] = df_display["Net price"].apply(eur_fmt)
    if "Total" in df_display:
        df_display["Total"] = df_display["Total"].apply(eur_fmt)

    st.dataframe(df_display, use_container_width=True, hide_index=True)

    c1, c2 = st.columns([1,2])
    with c1:
        if st.button("🗑️ Clear all positions", help="Reset the table"):
            reset_lines()
            st.warning("All positions cleared.")

    with c2:
        auto_filename = sanitize_filename(order_no or "supplier_order")

        def build_pdf(meta, lines_df):
            left_margin = 15*mm
            right_margin = 15*mm
            top_margin = 12*mm
            bottom_margin = 12*mm
            buffer = BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=left_margin, rightMargin=right_margin,
                                    topMargin=top_margin, bottomMargin=bottom_margin)
            styles = getSampleStyleSheet()
            styles.add(ParagraphStyle(name="Small", fontSize=7, leading=8.5))
            styles.add(ParagraphStyle(name="Normal8", fontSize=8, leading=10))
            styles.add(ParagraphStyle(name="Header", fontSize=12, leading=14, spaceAfter=4))

            rows = max(1, len(lines_df))
            body_font = 8 if rows <= 18 else (7 if rows <= 24 else 6)
            small_font = 7 if body_font >= 7 else 6

            story = []

            right_header = f"<b>Order</b><br/>Our Order No.: {meta.get('order_no','')}"
            if meta.get('your_order_ref'):
                right_header += f"<br/>Your order ref.: {meta.get('your_order_ref')}"
            right_header += f"<br/>VAT ID: {meta.get('vat_id','')}"

            title_table = Table([
                [
                    Paragraph("<b>ROTOGAL, S.L.U.</b><br/>POL. IND. ESPIÑERIA, PARC.36B<br/>15930 Boiro, A Coruña<br/>Spain", styles["Normal8"]),
                    Paragraph(right_header, styles["Normal8"]),
                ]
            ], colWidths=[100*mm, 70*mm])
            title_table.setStyle(TableStyle([
                ("VALIGN", (0,0), (-1,-1), "TOP"),
                ("BOTTOMPADDING", (0,0), (-1,-1), 2),
                ("TOPPADDING", (0,0), (-1,-1), 0),
            ]))
            story.append(title_table)
            story.append(Spacer(1, 3))

            shipping_html = clean(meta["ship_to"]).replace("\n","<br/>") if meta["ship_to"] else ""
            billing_html = clean(meta["bill_to"]).replace("\n","<br/>")
            meta_table = Table([
                [
                    Paragraph("<b>Shipping address:</b><br/>%s" % shipping_html, styles["Normal8"]),
                    Paragraph("<b>Billing address:</b><br/>%s" % billing_html, styles["Normal8"]),
                    Paragraph("Page: 1<br/>Date: %s<br/>Contact person: %s" % (meta["date_str"], meta["contact_person"]), styles["Normal8"]),
                ]
            ], colWidths=[65*mm, 65*mm, 40*mm])
            meta_table.setStyle(TableStyle([
                ("VALIGN", (0,0), (-1,-1), "TOP"),
                ("BOX", (0,0), (-1,-1), 0.25, colors.black),
                ("INNERGRID", (0,0), (-1,-1), 0.25, colors.black),
                ("BACKGROUND", (0,0), (-1,-1), colors.whitesmoke),
                ("LEFTPADDING", (0,0), (-1,-1), 2),
                ("RIGHTPADDING", (0,0), (-1,-1), 2),
                ("TOPPADDING", (0,0), (-1,-1), 1),
                ("BOTTOMPADDING", (0,0), (-1,-1), 1),
            ]))
            story.append(meta_table)
            story.append(Spacer(1, 4))

            base_mm = [10, 18, 85, 35, 12, 30, 30]
            target_total_mm = 180
            col_w_mm = scale_mm(base_mm, target_total_mm)
            col_w_pts = [w*mm for w in col_w_mm]

            header = ["Pos.", "Quantity", "Article", "Note", "VAT %", "Net price (EUR)", "Total (EUR)"]
            data = [header]
            pos = 1
            net_sum = 0.0
            for _, row in lines_df.iterrows():
                pg = clean(row.get('Product group',''))
                model = clean(row.get('Model',''))
                color = clean(row.get('Color',''))
                wall = clean(row.get('Wall build',''))
                drain = clean(row.get('Drain',''))
                note = clean(row.get('Note',''))

                parts = []
                if pg: parts.append(pg)
                if model: parts.append(f"Mod. {model}")
                if wall: parts.append(f"({wall})")
                if color: parts.append(color)
                if pg == "Bins" and drain and drain.lower() != "none":
                    parts.append(f"drain: {drain}")

                article = ", ".join(parts)
                net = float(row.get("Net price", 0.0) or 0.0)
                total = float(row.get("Total", 0.0) or 0.0)
                net_sum += total
                data.append([
                    str(pos),
                    str(int(row.get("Quantity", 0) or 0)),
                    Paragraph(article, ParagraphStyle(name="Cell", fontSize=body_font, leading=body_font+1)),
                    Paragraph(note, ParagraphStyle(name="Cell", fontSize=body_font, leading=body_font+1)),
                    f"{int(meta['vat_rate']*100)}%",
                    eur_fmt(net),
                    eur_fmt(total)
                ])
                pos += 1

            tbl = Table(data, colWidths=col_w_pts, repeatRows=1)
            tbl.setStyle(TableStyle([
                ("GRID", (0,0), (-1,-1), 0.25, colors.black),
                ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
                ("ALIGN", (0,0), (-1,0), "CENTER"),
                ("VALIGN", (0,0), (-1,-1), "TOP"),
                ("FONTSIZE", (0,0), (-1,0), body_font),
                ("FONTSIZE", (0,1), (-1,-1), body_font),
                ("LEFTPADDING", (0,0), (-1,-1), 2),
                ("RIGHTPADDING", (0,0), (-1,-1), 2),
                ("TOPPADDING", (0,0), (-1,-1), 1),
                ("BOTTOMPADDING", (0,0), (-1,-1), 1),
                ("ALIGN", (1,1), (1,-1), "RIGHT"),
                ("ALIGN", (5,1), (6,-1), "RIGHT"),
            ]))
            story.append(tbl)
            story.append(Spacer(1, 4))

            vat_amount = net_sum * meta["vat_rate"]
            gross = net_sum + vat_amount
            totals_table = Table([
                ["Net price:", eur_fmt(net_sum), "EUR"],
                [f"VAT ({int(meta['vat_rate']*100)}%):", eur_fmt(vat_amount), "EUR"],
                ["Gross price:", eur_fmt(gross), "EUR"],
            ], colWidths=[60*mm, 30*mm, 30*mm])
            totals_table.setStyle(TableStyle([
                ("ALIGN", (0,0), (-1,-1), "RIGHT"),
                ("FONTSIZE", (0,0), (-1,-1), body_font),
                ("LEFTPADDING", (0,0), (-1,-1), 2),
                ("RIGHTPADDING", (0,0), (-1,-1), 2),
            ]))
            story.append(totals_table)
            story.append(Spacer(1, 4))

            styles.add(ParagraphStyle(name="Tiny", fontSize=small_font, leading=small_font+1))
            story.append(Paragraph(
                "Customer protection, neutrality and on-time delivery are taken for granted. "
                "Please make sure to give Rotogal reference numbers with any query (invoice, delivery note). "
                "We kindly ask for a written confirmation of order.",
                styles["Tiny"]
            ))
            story.append(Spacer(1, 4))

            footer_right = f"VAT ID: {meta['vat_id']}\n" + meta["footer_right_extra"]
            footer_table = Table([
                [Paragraph(meta["footer_left"].replace("\n","<br/>"), styles["Tiny"]),
                 Paragraph(footer_right.replace("\n","<br/>"), styles["Tiny"])]
            ], colWidths=[90*mm, 90*mm])
            footer_table.setStyle(TableStyle([
                ("VALIGN", (0,0), (-1,-1), "TOP"),
                ("LEFTPADDING", (0,0), (-1,-1), 1),
                ("RIGHTPADDING", (0,0), (-1,-1), 1),
            ]))
            story.append(footer_table)

            doc.build(story)
            buffer.seek(0)
            return buffer

        meta = {
            "company": company,
            "contact_person": contact_person,
            "tel": tel,
            "email": email,
            "order_no": order_no,
            "your_order_ref": your_order_ref,
            "date_str": date_val.strftime("%d.%m.%Y") if date_val else "",
            "ship_to": ship_to,
            "bill_to": bill_to,
            "vat_id": vat_id,
            "vat_rate": vat_rate,
            "footer_left": footer_left,
            "footer_right_extra": footer_right_extra,
        }

        pdf_buffer = build_pdf(meta, pd.DataFrame(st.session_state.order_lines))

        # Microsoft Graph actions
        if (auto_upload or auto_email):
            if not all([tenant_id, client_id, client_secret, graph_user_upn]):
                st.warning("To use OneDrive/Email, please fill Tenant ID, Client ID, Client Secret and User UPN.")
            else:
                token, err = get_graph_token(tenant_id, client_id, client_secret)
                if token:
                    if auto_upload:
                        code, txt = onedrive_upload_file(token, graph_user_upn, onedrive_folder, auto_filename, pdf_buffer.getvalue())
                        if 200 <= code < 300:
                            st.success("Uploaded to OneDrive.")
                        else:
                            st.error(f"OneDrive upload failed ({code}): {txt}")
                    if auto_email and email_to.strip():
                        to_list = [e.strip() for e in email_to.split(",") if e.strip()]
                        subj = f"Order {order_no}"
                        body = f"<p>Hello,</p><p>Please find attached our order <b>{order_no}</b>.</p><p>Best regards,<br>{contact_person}</p>"
                        code, txt = graph_send_mail(token, graph_user_upn, to_list, subj, body, pdf_buffer.getvalue(), auto_filename)
                        if 200 <= code < 300:
                            st.success("Email sent via Microsoft Graph.")
                        else:
                            st.error(f"Email send failed ({code}): {txt}")
                else:
                    st.error(f"Graph auth failed: {err or 'Unknown error'}")

        st.download_button(
            label=f"📄 Download PDF ({auto_filename})",
            data=pdf_buffer,
            file_name=auto_filename,
            mime="application/pdf",
        )

st.divider()
st.caption("This build uses Microsoft Graph only (OneDrive upload + email). Fill your Tenant/Client/Secret/UPN in the sidebar or via Streamlit Secrets.")
