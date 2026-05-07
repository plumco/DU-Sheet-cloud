import streamlit as st
import pandas as pd
import io
import datetime
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import copy

# ─── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Huliot SO Form",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  .main-header {
    background: linear-gradient(90deg, #1a3a6b 0%, #2563b0 100%);
    padding: 18px 24px;
    border-radius: 10px;
    margin-bottom: 18px;
  }
  .main-header h1 { color: white; margin: 0; font-size: 1.6rem; }
  .main-header p  { color: #cce0ff; margin: 0; font-size: 0.88rem; }
  .section-card {
    background: #f8faff;
    border: 1px solid #dce8fb;
    border-radius: 8px;
    padding: 16px;
    margin-bottom: 12px;
  }
  .section-title {
    color: #1a3a6b;
    font-weight: 700;
    font-size: 0.95rem;
    border-bottom: 2px solid #2563b0;
    padding-bottom: 6px;
    margin-bottom: 12px;
  }
  .metric-box {
    background: #1a3a6b;
    color: white;
    border-radius: 8px;
    padding: 14px 18px;
    text-align: center;
  }
  .metric-box .val { font-size: 1.6rem; font-weight: 700; }
  .metric-box .lbl { font-size: 0.78rem; opacity: 0.8; }
  .stButton > button {
    background: #2563b0;
    color: white;
    border-radius: 6px;
    border: none;
    padding: 8px 20px;
  }
  .stButton > button:hover { background: #1a3a6b; }
  .product-row { background: #fff; border-radius: 6px; padding: 10px; margin: 4px 0; border: 1px solid #e2ecfb; }
  .remove-btn > button { background: #dc2626 !important; padding: 4px 12px !important; font-size: 0.8rem !important; }
  div[data-testid="stDataEditor"] { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# ─── Data Loaders ──────────────────────────────────────────────────────────────
EXCEL_FILE = Path(__file__).parent / "SO_FORMAT_PROJECT_-_APRIL_2026.xlsx"

@st.cache_data
def load_product_master():
    df = pd.read_excel(EXCEL_FILE, sheet_name="Product Master", header=0)
    df.columns = ["item_code","description","dn","moq","list_price","category","sub_group","hsn_code","gst_rate","cat2"]
    df = df.dropna(subset=["item_code"])
    df["search_label"] = df["item_code"].astype(str) + " │ " + df["description"].astype(str) + " │ DN" + df["dn"].astype(str)
    return df

@st.cache_data
def load_validations():
    df = pd.read_excel(EXCEL_FILE, sheet_name="Data Validations", header=0)
    return df

def get_list(df, col):
    return [str(v) for v in df[col].dropna().tolist()]

products_df = load_product_master()
val_df      = load_validations()

sales_persons   = get_list(val_df, "Sales Person")
payment_terms   = get_list(val_df, "Payment Terms")
regions         = get_list(val_df, "Region")
project_types   = get_list(val_df, "Project Type")
customer_types  = get_list(val_df, "Customer Type")
freight_terms   = get_list(val_df, "Freight Terms")
order_validities = get_list(val_df, "Order Validity")
rera_types      = get_list(val_df, "RERA")
cd_options      = sorted([round(v, 2) for v in val_df["CD %"].dropna().tolist()])

# ─── Session State ─────────────────────────────────────────────────────────────
def init_state():
    defaults = {
        "line_items": [],
        "so_no": "",
        "po_no": "",
        "po_date": datetime.date.today(),
        "order_date": datetime.date.today(),
        "sales_person": sales_persons[2] if len(sales_persons) > 2 else "",
        "business_type": "Project",
        "business_segment": "Housing-Residential Project - HR",
        "customer_po_no": "",
        "bitrix_id": "",
        "customer_type": "Builder",
        "rera_no": "",
        "order_validity": "7 Days",
        "other_info": "",
        "cash_discount": 0.03,
        "freight": "FOR",
        "payment_term": "Advance",
        "bill_name": "",
        "bill_address": "",
        "bill_city": "",
        "bill_pin": "",
        "bill_gstin": "",
        "ship_name": "",
        "ship_address": "",
        "ship_city": "",
        "ship_pin": "",
        "ship_gstin": "",
        "consultant_name": "",
        "consultant_contact": "",
        "developer_name": "",
        "site_name": "",
        "other_remark": "",
        "contact_person": "",
        "contact_mobile": "",
        "search_query": "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ─── Header ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>📋 Huliot Pipes & Fittings — Sales Order Form</h1>
  <p>Format No: HPF-01 &nbsp;|&nbsp; Rev. 01 &nbsp;|&nbsp; Issue Date: 10.04.2026</p>
</div>
""", unsafe_allow_html=True)

# ─── Tabs ─────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📝 Order Details", "🔍 Add Products", "📥 Summary & Download"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — ORDER DETAILS
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    # ── Row A: Order Identity
    st.markdown('<div class="section-title">📌 Order Identity</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.session_state.so_no        = st.text_input("Sales Order No", value=st.session_state.so_no)
    with c2:
        st.session_state.po_no        = st.text_input("PO No", value=st.session_state.po_no)
    with c3:
        st.session_state.po_date      = st.date_input("PO Date", value=st.session_state.po_date)
    with c4:
        st.session_state.order_date   = st.date_input("Order Date", value=st.session_state.order_date)

    st.markdown("---")

    # ── Row B: Sales & Project
    st.markdown('<div class="section-title">👤 Sales & Project Info</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        idx = sales_persons.index(st.session_state.sales_person) if st.session_state.sales_person in sales_persons else 0
        st.session_state.sales_person     = st.selectbox("Sales Person", sales_persons, index=idx)
    with c2:
        bt_opts = ["Project", "Retail"]
        bt_idx  = bt_opts.index(st.session_state.business_type) if st.session_state.business_type in bt_opts else 0
        st.session_state.business_type    = st.selectbox("Business Type", bt_opts, index=bt_idx)
    with c3:
        seg_idx = project_types.index(st.session_state.business_segment) if st.session_state.business_segment in project_types else 1
        st.session_state.business_segment = st.selectbox("Business Segment / Project Type", project_types, index=seg_idx)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.session_state.rera_no          = st.text_input("RERA Project No *", value=st.session_state.rera_no,
                                                           help="Mandatory — order won't process without valid RERA")
    with c2:
        ct_idx = customer_types.index(st.session_state.customer_type) if st.session_state.customer_type in customer_types else 1
        st.session_state.customer_type    = st.selectbox("Customer Type", customer_types, index=ct_idx)
    with c3:
        st.session_state.customer_po_no   = st.text_input("Customer PO No", value=st.session_state.customer_po_no)
    with c4:
        st.session_state.bitrix_id        = st.text_input("Bitrix Deal ID", value=st.session_state.bitrix_id)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.session_state.contact_person   = st.text_input("Contact Person Name", value=st.session_state.contact_person)
    with c2:
        st.session_state.contact_mobile   = st.text_input("Contact Mobile", value=st.session_state.contact_mobile)
    with c3:
        ov_idx = order_validities.index(st.session_state.order_validity) if st.session_state.order_validity in order_validities else 0
        st.session_state.order_validity   = st.selectbox("Order Validity", order_validities, index=ov_idx)

    st.markdown("---")

    # ── Row C: Commercial Terms
    st.markdown('<div class="section-title">💰 Commercial Terms</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        cd_idx = cd_options.index(st.session_state.cash_discount) if st.session_state.cash_discount in cd_options else 1
        st.session_state.cash_discount    = st.selectbox("Cash Discount %", cd_options, index=cd_idx,
                                                          format_func=lambda x: f"{x*100:.0f}%")
    with c2:
        ft_idx = freight_terms.index(st.session_state.freight) if st.session_state.freight in freight_terms else 0
        st.session_state.freight          = st.selectbox("Freight / Term of Delivery", freight_terms, index=ft_idx)
    with c3:
        pt_idx = payment_terms.index(st.session_state.payment_term) if st.session_state.payment_term in payment_terms else 0
        st.session_state.payment_term     = st.selectbox("Payment Term", payment_terms, index=pt_idx)

    st.markdown("---")

    # ── Row D: Bill To / Ship To
    st.markdown('<div class="section-title">🏢 Billing & Shipping</div>', unsafe_allow_html=True)
    col_bill, col_ship = st.columns(2)
    with col_bill:
        st.markdown("**Receiver (Bill To)**")
        st.session_state.bill_name    = st.text_input("Name",    value=st.session_state.bill_name,    key="bn")
        st.session_state.bill_address = st.text_area("Address", value=st.session_state.bill_address, key="ba", height=70)
        c1, c2, c3 = st.columns(3)
        with c1: st.session_state.bill_city  = st.text_input("City",     value=st.session_state.bill_city,  key="bc")
        with c2: st.session_state.bill_pin   = st.text_input("Pin Code", value=st.session_state.bill_pin,   key="bp")
        with c3: st.session_state.bill_gstin = st.text_input("GSTIN",    value=st.session_state.bill_gstin, key="bg")
    with col_ship:
        st.markdown("**Consignee (Ship To)**")
        st.session_state.ship_name    = st.text_input("Name",    value=st.session_state.ship_name,    key="sn")
        st.session_state.ship_address = st.text_area("Address", value=st.session_state.ship_address, key="sa", height=70)
        c1, c2, c3 = st.columns(3)
        with c1: st.session_state.ship_city  = st.text_input("City",     value=st.session_state.ship_city,  key="sc")
        with c2: st.session_state.ship_pin   = st.text_input("Pin Code", value=st.session_state.ship_pin,   key="sp")
        with c3: st.session_state.ship_gstin = st.text_input("GSTIN",    value=st.session_state.ship_gstin, key="sg")

    st.markdown("---")

    # ── Row E: Additional Info
    st.markdown('<div class="section-title">📎 Additional Info</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.session_state.consultant_name    = st.text_input("Consultant / Architect Firm", value=st.session_state.consultant_name)
        st.session_state.developer_name     = st.text_input("Developer / Bungalow Owner",  value=st.session_state.developer_name)
    with c2:
        st.session_state.consultant_contact = st.text_input("Consultant Contact No.", value=st.session_state.consultant_contact)
        st.session_state.site_name          = st.text_input("Site / Project Name",    value=st.session_state.site_name)
    st.session_state.other_remark           = st.text_area("Other Remark", value=st.session_state.other_remark, height=60)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — ADD PRODUCTS
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-title">🔍 Search & Add Products</div>', unsafe_allow_html=True)

    # ── Search bar
    search_q = st.text_input("Search by Item Code or Description", placeholder="e.g. 5751 or ULTRA SILENT or 110mm elbow",
                              value=st.session_state.search_query, key="search_input")
    st.session_state.search_query = search_q

    if search_q:
        mask = (
            products_df["item_code"].str.contains(search_q, case=False, na=False) |
            products_df["description"].str.contains(search_q, case=False, na=False) |
            products_df["dn"].astype(str).str.contains(search_q, case=False, na=False)
        )
        filtered = products_df[mask].head(60)
    else:
        filtered = products_df.head(0)

    if not filtered.empty:
        st.caption(f"Found {len(filtered)} products")

        # Show table header
        hdr = st.columns([3, 2, 1, 1, 1, 1, 1.2])
        for h, t in zip(hdr, ["Description", "Category", "DN", "MOQ", "List Price ₹", "Disc %", ""]):
            h.markdown(f"**{t}**")

        for _, row in filtered.iterrows():
            cols = st.columns([3, 2, 1, 1, 1, 1, 1.2])
            cols[0].markdown(f"<small>{row['item_code']}<br>{row['description']}</small>", unsafe_allow_html=True)
            cols[1].markdown(f"<small>{row['category']}</small>", unsafe_allow_html=True)
            cols[2].write(str(row['dn']))
            cols[3].write(str(int(row['moq'])))
            cols[4].write(f"₹{row['list_price']:,.0f}")
            disc_key = f"disc_{row['item_code']}"
            if disc_key not in st.session_state:
                st.session_state[disc_key] = 0.52
            disc_val = cols[5].number_input("", min_value=0.0, max_value=0.99, step=0.01,
                                             value=st.session_state[disc_key], key=f"di_{row['item_code']}",
                                             label_visibility="collapsed", format="%.2f")
            st.session_state[disc_key] = disc_val
            if cols[6].button("➕ Add", key=f"add_{row['item_code']}"):
                # Check if already in list
                existing = [i for i, item in enumerate(st.session_state.line_items) if item["item_code"] == row["item_code"]]
                if existing:
                    st.session_state.line_items[existing[0]]["qty"] += int(row['moq'])
                    st.toast(f"Qty updated for {row['item_code']}", icon="✅")
                else:
                    cd  = st.session_state.cash_discount
                    net_disc = 1 - (1 - disc_val) * (1 - cd)
                    net_price = round(row['list_price'] * (1 - net_disc), 2)
                    st.session_state.line_items.append({
                        "item_code":   row["item_code"],
                        "description": row["description"],
                        "dn":          row["dn"],
                        "category":    row["category"],
                        "sub_group":   row["sub_group"],
                        "moq":         int(row["moq"]),
                        "qty":         int(row["moq"]),
                        "hsn_code":    row["hsn_code"],
                        "list_price":  row["list_price"],
                        "disc_pct":    disc_val,
                        "cash_disc":   cd,
                        "net_disc":    round(net_disc, 4),
                        "net_price":   net_price,
                    })
                    st.toast(f"Added {row['item_code']}", icon="✅")
                st.rerun()
    elif search_q:
        st.info("No products found. Try different keywords.")

    st.markdown("---")

    # ── Line Items Table
    st.markdown('<div class="section-title">🛒 Line Items</div>', unsafe_allow_html=True)

    if not st.session_state.line_items:
        st.info("No products added yet. Search above and click ➕ Add.")
    else:
        # Column headers
        h = st.columns([0.4, 2.8, 0.8, 1.2, 0.8, 1.0, 1.0, 1.2, 0.5])
        for col, lbl in zip(h, ["S.No", "Description", "DN", "Category", "MOQ", "Qty", "List ₹", "Net Price ₹", ""]):
            col.markdown(f"**{lbl}**")

        to_remove = []
        for idx, item in enumerate(st.session_state.line_items):
            cols = st.columns([0.4, 2.8, 0.8, 1.2, 0.8, 1.0, 1.0, 1.2, 0.5])
            cols[0].write(str(idx + 1))
            cols[1].markdown(f"<small><b>{item['item_code']}</b><br>{item['description']}</small>", unsafe_allow_html=True)
            cols[2].write(str(item['dn']))
            cols[3].markdown(f"<small>{item['category']}</small>", unsafe_allow_html=True)
            cols[4].write(str(item['moq']))

            new_qty = cols[5].number_input("qty", min_value=0, step=1, value=item['qty'],
                                            key=f"qty_{idx}", label_visibility="collapsed")
            if new_qty != item['qty']:
                st.session_state.line_items[idx]['qty'] = new_qty
                # Recalc
                cd       = st.session_state.cash_discount
                disc     = item['disc_pct']
                net_disc = round(1 - (1 - disc) * (1 - cd), 4)
                st.session_state.line_items[idx]['cash_disc'] = cd
                st.session_state.line_items[idx]['net_disc']  = net_disc
                st.session_state.line_items[idx]['net_price'] = round(item['list_price'] * (1 - net_disc), 2)

            cols[6].write(f"₹{item['list_price']:,.0f}")
            cols[7].write(f"₹{item['net_price']:,.2f}")

            if cols[8].button("🗑", key=f"rem_{idx}"):
                to_remove.append(idx)

        for i in reversed(to_remove):
            st.session_state.line_items.pop(i)
        if to_remove:
            st.rerun()

        # Summary metrics
        total_basic = sum(item['qty'] * item['net_price'] for item in st.session_state.line_items)
        total_items = len(st.session_state.line_items)
        total_qty   = sum(item['qty'] for item in st.session_state.line_items)

        st.markdown("---")
        m1, m2, m3 = st.columns(3)
        m1.markdown(f'<div class="metric-box"><div class="val">{total_items}</div><div class="lbl">Products</div></div>', unsafe_allow_html=True)
        m2.markdown(f'<div class="metric-box"><div class="val">{total_qty}</div><div class="lbl">Total Qty</div></div>', unsafe_allow_html=True)
        m3.markdown(f'<div class="metric-box"><div class="val">₹{total_basic:,.0f}</div><div class="lbl">Order Basic Value</div></div>', unsafe_allow_html=True)

        if st.button("🗑️ Clear All Line Items"):
            st.session_state.line_items = []
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — SUMMARY & DOWNLOAD
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    s = st.session_state
    total_basic = sum(item['qty'] * item['net_price'] for item in s.line_items)

    # Validation
    errors = []
    if not s.so_no:            errors.append("Sales Order No is blank")
    if not s.sales_person:     errors.append("Sales Person not selected")
    if not s.rera_no:          errors.append("RERA Project No is mandatory")
    if not s.bill_name:        errors.append("Bill-to Name is blank")
    if not s.line_items:       errors.append("No line items added")

    if errors:
        for e in errors:
            st.warning(f"⚠️ {e}")
    else:
        st.success("✅ All required fields filled. Ready to download.")

    st.markdown("---")
    st.markdown('<div class="section-title">📋 Order Summary</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"""
| Field | Value |
|---|---|
| **SO No** | {s.so_no or '—'} |
| **Date** | {s.order_date} |
| **Sales Person** | {s.sales_person} |
| **Business Segment** | {s.business_segment} |
| **Customer Type** | {s.customer_type} |
| **RERA No** | {s.rera_no or '—'} |
""")
    with c2:
        st.markdown(f"""
| Field | Value |
|---|---|
| **Cash Discount** | {s.cash_discount*100:.0f}% |
| **Payment Term** | {s.payment_term} |
| **Freight** | {s.freight} |
| **Order Validity** | {s.order_validity} |
| **Bill To** | {s.bill_name or '—'} |
| **Order Value** | ₹{total_basic:,.2f} |
""")

    if s.line_items:
        st.markdown("**Line Items:**")
        rows = []
        for i, item in enumerate(s.line_items):
            rows.append({
                "S.No": i+1,
                "Item Code": item["item_code"],
                "Description": item["description"],
                "DN": item["dn"],
                "Qty": item["qty"],
                "List Price ₹": f"₹{item['list_price']:,.0f}",
                "Disc %": f"{item['disc_pct']*100:.1f}%",
                "Net Price ₹": f"₹{item['net_price']:,.2f}",
                "Basic Value ₹": f"₹{item['qty']*item['net_price']:,.2f}",
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown('<div class="section-title">📥 Download</div>', unsafe_allow_html=True)

    def build_excel():
        wb = load_workbook(EXCEL_FILE)
        ws = wb["Sales Order"]

        def w(cell, val):
            ws[cell] = val

        # Header fields
        w("C6",  s.business_type)
        w("F6",  s.sales_person)
        w("J6",  s.po_no)
        w("N6",  s.po_date.strftime("%d.%m.%Y"))

        w("C7",  s.business_segment)
        w("F7",  s.rera_no)
        w("J7",  s.order_validity)

        w("C8",  s.customer_po_no)
        w("J8",  s.so_no)
        w("N8",  s.order_date.strftime("%d.%m.%Y"))

        w("C9",  s.bitrix_id)
        w("F9",  s.customer_type)
        w("J9",  s.contact_person)
        w("N9",  s.contact_mobile)

        # Bill to
        w("B11", s.bill_name)
        w("B12", s.bill_address)
        w("B14", s.bill_city)
        w("B15", s.bill_pin)
        w("B16", s.bill_gstin)

        # Ship to
        w("I11", s.ship_name)
        w("I12", s.ship_address)
        w("I14", s.ship_city)
        w("I15", s.ship_pin)
        w("I16", s.ship_gstin)

        w("C17", s.cash_discount)
        w("C18", s.freight)
        w("C19", s.payment_term)

        w("D20", s.consultant_name)
        w("L20", s.consultant_contact)
        w("D21", s.developer_name)
        w("L21", s.site_name)
        w("C22", s.other_remark)

        # Clear existing line item rows (25 to 102)
        for r in range(25, 103):
            for col in ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O"]:
                ws[f"{col}{r}"] = None

        # Write line items
        for i, item in enumerate(s.line_items):
            r = 25 + i
            if r > 102:
                break
            net_disc = round(1 - (1 - item["disc_pct"]) * (1 - s.cash_discount), 4)
            net_price = round(item["list_price"] * (1 - net_disc), 2)
            ws[f"A{r}"] = i + 1
            ws[f"B{r}"] = item["item_code"]
            ws[f"C{r}"] = item["description"]
            ws[f"D{r}"] = item["dn"]
            ws[f"E{r}"] = item["category"]
            ws[f"F{r}"] = item["moq"]
            ws[f"G{r}"] = item["qty"]
            ws[f"H{r}"] = item["hsn_code"]
            ws[f"I{r}"] = item["list_price"]
            ws[f"J{r}"] = item["disc_pct"]
            ws[f"K{r}"] = s.cash_discount
            ws[f"L{r}"] = net_disc
            ws[f"M{r}"] = net_price
            ws[f"N{r}"] = item["qty"] * net_price
            ws[f"O{r}"] = item["sub_group"]

        # Update Dashboard sheet
        try:
            wd = wb["Dash Board"]
            wd["B5"]  = s.so_no
            wd["B6"]  = s.sales_person
            wd["B7"]  = s.bill_name
            wd["B8"]  = s.bill_name
            wd["B9"]  = s.bill_name
            wd["B10"] = s.bitrix_id
            wd["B12"] = s.cash_discount
            wd["B14"] = s.freight
            wd["B15"] = s.payment_term
            wd["B16"] = s.rera_no
        except Exception:
            pass

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    col_dl, col_info = st.columns([2, 3])
    with col_dl:
        if st.button("🔄 Generate Excel File", type="primary", disabled=bool(errors)):
            with st.spinner("Building Excel..."):
                try:
                    excel_bytes = build_excel()
                    fname = f"SO_{s.so_no or 'DRAFT'}_{s.order_date.strftime('%d%m%Y')}.xlsx"
                    st.download_button(
                        label="📥 Download Excel (.xlsx)",
                        data=excel_bytes,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except Exception as ex:
                    st.error(f"Error: {ex}")
    with col_info:
        st.info("💡 Downloaded file uses the original Huliot SO template format with your data pre-filled. Open in Excel to verify formulas.")

    if errors:
        st.warning("Fix the warnings above before generating Excel.")
