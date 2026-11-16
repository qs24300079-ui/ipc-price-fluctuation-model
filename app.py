"""
IPC Resource Composition + Price Fluctuation Calculator  –  Streamlit app
Run:  streamlit run app.py
"""
import io
from pathlib import Path

import pandas as pd
import streamlit as st

from ipc_resource_engine import (DEFAULT_WORKBOOK, IPCModel, merge_results, price_fluctuation)
from pf_model import PFCalculator, PX_TOTAL

st.set_page_config(page_title="IPC → Resource Composition", layout="wide")
st.title("IPC → Resource Composition Calculator")
st.caption("Enter an IPC value or the work completed. The tool shows which resources were consumed "
           "and how much, so price fluctuation can be applied to the right values.")


@st.cache_resource(show_spinner="Reading workbook …")
def load_model(source):
    src = (lambda: io.BytesIO(source)) if isinstance(source, bytes) else (lambda: source)
    model = IPCModel.from_workbook(src())
    return model, PFCalculator(model, src())


with st.sidebar:
    st.header("Model workbook")
    up = st.file_uploader("Upload the framework workbook (.xlsx)", type="xlsx")
    if up is not None:
        model, calc = load_model(up.getvalue())
    elif Path(DEFAULT_WORKBOOK).exists():
        model, calc = load_model(DEFAULT_WORKBOOK)
        st.success(f"Using {DEFAULT_WORKBOOK}")
    else:
        st.error("Upload the workbook to begin.")
        st.stop()
    st.write(f"Profile: **{model.profile.name}**  ·  PF sheets: {len(calc.sheets)}")
    st.write(f"{len(model.groups)} work groups · {len(model.resources)} resources · "
             f"IPCs in workbook: {min(model.available_ipcs)}–{max(model.available_ipcs)}")

p = model.profile
if "session" not in st.session_state:
    st.session_state.session = []


def show_result(res, key):
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Input value (LKR)", f"{res.input_value:,.0f}")
    c2.metric("Adjustable base (LKR)", f"{res.adjustable_base:,.0f}")
    c3.metric("Non-adj. + disregarded (LKR)", f"{res.non_adjustable + res.disregarded:,.0f}")
    c4.metric("Resources mapped (LKR)", f"{res.total_resources:,.0f}")
    for n in res.notes:
        st.caption(n)

    t = res.nonzero()
    left, right = st.columns([3, 2])
    with left:
        st.subheader("Resource values")
        st.dataframe(
            t.rename(columns={"code": "Code", "name": "Resource", "type": "Type",
                              "amount_lkr": "Amount (LKR)", "share_of_adjustable_pct": "% of adjustable"}),
            hide_index=True, use_container_width=True,
            column_config={"Amount (LKR)": st.column_config.NumberColumn(format="%,.2f"),
                           "% of adjustable": st.column_config.NumberColumn(format="%.2f")})
    with right:
        st.subheader("Composition")
        st.bar_chart(t.set_index("code")["amount_lkr"])
        st.write(res.by_type().apply(lambda v: f"{v:,.0f}").rename("LKR by type"))

    with st.expander("Which work groups produced these resources?"):
        st.dataframe(res.by_group.round(2), use_container_width=True)

    with st.expander("Price fluctuation on these resources (optional)"):
        st.write("Enter base-date and current price indices (or prices) per resource. "
                 "Fluctuation = amount × (current ÷ base − 1).")
        seed = pd.DataFrame({"code": t["code"], "base_index": 100.0, "current_index": 100.0})
        ed = st.data_editor(seed, hide_index=True, key=f"idx_{key}", use_container_width=True, disabled=["code"])
        # Keep values numeric even when a browser/editor returns edited cells as text.
        ed["base_index"] = pd.to_numeric(ed["base_index"], errors="coerce")
        ed["current_index"] = pd.to_numeric(ed["current_index"], errors="coerce")
        pf = price_fluctuation(res, dict(zip(ed.code, ed.base_index)), dict(zip(ed.code, ed.current_index)))
        st.dataframe(pf[["code", "name", "amount_lkr", "base_index", "current_index", "change_pct", "fluctuation_lkr"]],
                     hide_index=True, use_container_width=True)
        st.metric("Total price fluctuation (LKR)", f"{pf['fluctuation_lkr'].sum():,.2f}")

    buf = io.BytesIO()
    res.to_excel(buf)
    st.download_button("⬇ Download results (Excel)", buf.getvalue(), file_name="ipc_resource_result.xlsx", key=f"dl_{key}")


tab1, tab2, tab3, tab4, tab5 = st.tabs(["1 · IPC value", "2 · Work completed by group",
                                        "3 · Work completed by BOQ row", "4 · Running total", "5 · Price fluctuation"])

with tab1:
    st.write("Use this when only the IPC amount is known.")
    a, b = st.columns(2)
    ipcs = model.available_ipcs
    ipc_no = a.number_input("IPC number", min_value=1, value=int(ipcs[min(5, len(ipcs) - 1)]), step=1)
    value = b.number_input("IPC value this period (LKR)", min_value=0.0,
                           value=float(model.targets.get(int(ipc_no), 10_000_000.0)), step=100_000.0, format="%.2f")
    if int(ipc_no) not in ipcs:
        st.info("This IPC is not in the workbook – a uniform BOQ pro-rata profile is used.")
    try:
        show_result(model.resources_from_ipc(int(ipc_no), value), "t1")
    except Exception as e:
        st.error(str(e))

with tab2:
    st.write("Enter the value of work completed in each work group (LKR).")
    df = pd.DataFrame({"Work group": model.groups, "Work done (LKR)": 0.0})
    ed = st.data_editor(df, hide_index=True, disabled=["Work group"], use_container_width=True, height=350, key="group_editor")
    amounts = {r["Work group"]: r["Work done (LKR)"] for _, r in ed.iterrows() if r["Work done (LKR)"]}
    if amounts:
        r2 = model.resources_from_groups(amounts)
        show_result(r2, "t2")
        if st.button("➕ Add to running total", key="add2"):
            st.session_state.session.append(r2)
            st.toast("Added")
    else:
        st.info("Type at least one amount.")

with tab3:
    st.write("Enter the value completed against individual BOQ rows. Items split into components "
             "(e.g. lintels) have one row per component.")
    rows = model.items[[p.id_col, p.item_col, p.desc_col, p.group_col, p.treat_col, p.amount_col]].copy()
    rows[p.desc_col] = rows[p.desc_col].astype(str).str[:60]
    rows["Work done (LKR)"] = 0.0
    q = st.text_input("Filter (item no. or words)", "")
    if q:
        rows = rows[rows.apply(lambda r: q.lower() in " ".join(map(str, r.values)).lower(), axis=1)]
    ed3 = st.data_editor(rows, hide_index=True, height=350, key="row_editor", use_container_width=True,
                         disabled=[c for c in rows.columns if c != "Work done (LKR)"])
    amt3 = {r[p.id_col]: r["Work done (LKR)"] for _, r in ed3.iterrows() if r["Work done (LKR)"]}
    if amt3:
        r3 = model.resources_from_rows(amt3)
        show_result(r3, "t3")
        if st.button("➕ Add to running total", key="add3"):
            st.session_state.session.append(r3)
            st.toast("Added")
    else:
        st.info("Type at least one amount.")

with tab4:
    st.write(f"{len(st.session_state.session)} entries added.")
    if st.session_state.session:
        show_result(merge_results(st.session_state.session), "t4")
        if st.button("Clear running total"):
            st.session_state.session = []
            st.rerun()


with tab5:
    st.write("Price adjustment = factor × (V − Vna)/100 × Σ Px (Ixc − Ixb)/Ixb. "
             "Px comes from the resource composition of the IPC, so no waiting for the project to finish.")
    part = st.radio("What do you want to do?", ["Check a Price Fluctuation sheet of the workbook", "New valuation – any IPC value and indices"],
                    horizontal=True)

    def show_pf(r):
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("V − Vna (Rs.)", f"{r.v_vna:,.0f}")
        c2.metric("Factor", f"{r.factor:.2f}")
        c3.metric("Σ Px·ΔI/I", f"{r.sigma:.4f}")
        c4.metric("Price adjustment (Rs.)", f"{r.amount:,.0f}")
        for n in r.notes:
            st.caption(n)
        t = r.table[r.table.px > 0].sort_values("contribution")
        st.dataframe(t.rename(columns={"code": "Code", "name": "Resource", "px": "Px %", "ixb": "Ixb", "ixc": "Ixc",
                                       "change": "Change", "contribution": "Px × change"}),
                     hide_index=True, use_container_width=True,
                     column_config={"Change": st.column_config.NumberColumn(format="%.2f%%"),
                                    "Px %": st.column_config.NumberColumn(format="%.4f"),
                                    "Px × change": st.column_config.NumberColumn(format="%.4f")})

    if part.startswith("Check"):
        if not calc.sheets:
            st.info("This workbook has no 'Price Fluctuation N' sheets.")
        else:
            n = st.selectbox("Price Fluctuation No.", list(calc.sheets))
            s = calc.sheets[n]
            mode = st.radio("Calculation basis", ["As in the workbook sheet", "Model-driven (Px of the IPC valued, V − Vna from the model)"], horizontal=True)
            if mode.startswith("As"):
                r = calc.as_sheet(n)
            else:
                off = st.radio("Px of …", ["IPC valued (N + 1)", "IPC N (as in the sheets)"], horizontal=True)
                r = calc.model_driven(n, 1 if off.startswith("IPC valued") else 0)
            show_pf(r)
            st.caption(f"Workbook sheet shows: Σ = {s.sigma_sheet:.4f},  V − Vna = {s.vvna_sheet:,.2f},  price adjustment = {s.amount_sheet:,.2f}")
        with st.expander("All PF sheets: Excel vs Python, and the model-driven alternative"):
            if calc.sheets:
                st.dataframe(calc.compare_all().round(2), hide_index=True, use_container_width=True)
        with st.expander("Consistency checks of the PF sheets"):
            if calc.sheets:
                st.write("Cumulative non-adjustable figure (d) must not fall; V − Vna cannot exceed the work valued + change in material at site:")
                st.dataframe(calc.non_adjustable_check().round(0), hide_index=True, use_container_width=True)
    else:
        a, b, c = st.columns(3)
        ipc = a.number_input("IPC number", min_value=1, value=int(model.available_ipcs[min(5, len(model.available_ipcs) - 1)]), step=1, key="pf_ipc")
        val = b.number_input("IPC value this period (LKR)", min_value=0.0, value=float(model.targets.get(int(ipc), 10_000_000.0)),
                             step=100_000.0, format="%.2f", key="pf_val")
        m_prev = c.number_input("Material at site – previous (LKR)", min_value=0.0, value=0.0, step=10_000.0, key="pf_mp")
        m_cur = c.number_input("Material at site – current (LKR)", min_value=0.0, value=0.0, step=10_000.0, key="pf_mc")
        codes = list(model.resources.code)
        start = None
        if calc.sheets:
            src = st.selectbox("Start the indices from PF sheet", ["(all 100)"] + [f"PF {k}" for k in calc.sheets])
            if src != "(all 100)":
                start = calc.sheets[int(src.split()[1])].indices
        seed = pd.DataFrame({"code": codes,
                             "base_index": [float(start.ixb.get(k, 100.0)) if start is not None else 100.0 for k in codes],
                             "current_index": [float(start.ixc.get(k, 100.0)) if start is not None else 100.0 for k in codes]})
        ed = st.data_editor(seed, hide_index=True, disabled=["code"], use_container_width=True, height=300, key="pf_idx")
        r = calc.from_ipc(int(ipc), val, pd.Series(ed.base_index.values, index=ed.code), pd.Series(ed.current_index.values, index=ed.code),
                          m_prev, m_cur)
        show_pf(r)
