"""
Price-fluctuation module (ICTAD-type formula) on top of the IPC resource engine
===============================================================================
    Price adjustment = factor x (V - Vna) / 100 x  SUM_x  Px x (Ixc - Ixb) / Ixb

    Px   = percentage input of resource x for the work being valued. In the framework it comes
           from the resource composition of the IPC (Px scaled to sum to `px_total`, 96 here).
    V - Vna = value of ADJUSTABLE work in the valuation period (+ change in material at site)
    Ixb / Ixc = base-date / current-date price index of resource x
    factor = px_total / 100 (0.96 here)

Three ways to use it
--------------------
1. as_sheet(n)      reproduces a 'Price Fluctuation n' sheet of the workbook (Px of IPC n, V-Vna from the
                    sheet's cumulative figures) - used to validate Excel against Python
2. model_driven(n)  Px of the IPC actually valued (n + 1) and V - Vna taken from the model
3. from_ipc(k, value, indices)   any IPC value + any indices - no PF sheet needed

PF No. N values the work between cumulative IPC N and IPC N+1 (model IPC numbering).
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook

from ipc_resource_engine import IPCModel, Result

PX_TOTAL = {"validation": 96.0, "project1": 93.0}      # Px sum per project profile (factor = px_total / 100)
PF_PREFIX = "Price Fluctuation "


# ------------------------------------------------------------------ data classes
@dataclass
class PFSheet:
    number: int
    title: str
    prev: dict                    # date, b (cum. value of work), c (material at site), d (cum. non-adjustable)
    cur: dict
    indices: pd.DataFrame         # index code; columns name, px_sheet, ixb, ixc
    sigma_sheet: float
    vvna_sheet: float
    factor_sheet: float
    amount_sheet: float


@dataclass
class PFResult:
    label: str
    px: pd.Series
    table: pd.DataFrame           # code, name, px, ixb, ixc, change, contribution
    sigma: float
    v_vna: float
    factor: float
    amount: float
    notes: list = field(default_factory=list)

    def summary_text(self) -> str:
        lines = [f"{self.label}", f"  V - Vna                 : {self.v_vna:>16,.2f}",
                 f"  factor                  : {self.factor:>16.4f}",
                 f"  SUM Px(Ixc-Ixb)/Ixb     : {self.sigma:>16.4f}",
                 f"  PRICE ADJUSTMENT (Rs.)  : {self.amount:>16,.2f}", ""]
        t = self.table[self.table.px > 0].sort_values("contribution")
        lines.append(f"{'Code':<6}{'Resource':<34}{'Px %':>8}{'Ixb':>9}{'Ixc':>9}{'change':>9}{'Px x change':>12}")
        for _, r in t.iterrows():
            lines.append(f"{r.code:<6}{str(r['name'])[:32]:<34}{r.px:>8.4f}{r.ixb:>9.1f}{r.ixc:>9.1f}{r.change * 100:>8.2f}%{r.contribution:>12.4f}")
        return "\n".join(lines + [""] + [f"Note: {n}" for n in self.notes])


# --------------------------------------------------------------- reading PF sheets
def _num(v, default=0.0):
    return float(v) if isinstance(v, (int, float)) else default


def read_pf_sheets(path) -> dict:
    """Reads every 'Price Fluctuation N' sheet (indices + valuation data) from the workbook (path or uploaded file)."""
    import io
    if hasattr(path, "read"):
        data = path.read()
        wv, wf = load_workbook(io.BytesIO(data), data_only=True), load_workbook(io.BytesIO(data))
    else:
        wv, wf = load_workbook(str(path), data_only=True), load_workbook(str(path))
    out = {}
    for name in wv.sheetnames:
        m = re.fullmatch(PF_PREFIX + r"(\d+)", name)
        if not m:
            continue
        s, sf = wv[name], wf[name]
        rows, k = [], 2
        while s.cell(k, 1).value not in (None, "") and isinstance(s.cell(k, 4).value, (int, float)):
            rows.append([str(s.cell(k, 1).value).strip(), s.cell(k, 2).value, _num(s.cell(k, 3).value),
                         _num(s.cell(k, 4).value), _num(s.cell(k, 5).value)])
            k += 1
        idx = pd.DataFrame(rows, columns=["code", "name", "px_sheet", "ixb", "ixc"]).set_index("code")
        blocks = {"prev": {}, "cur": {}}
        section, title, vvna, pa, factor = None, "", np.nan, np.nan, 0.96
        for r in range(k, s.max_row + 1):
            t = str(s.cell(r, 2).value or "")
            if t.startswith("Price adjustment computation"): title = t
            elif t.startswith("Previous Valuation"): section = "prev"
            elif t.startswith("Current Valuation"): section = "cur"
            elif t == "Date" and section: blocks[section]["date"] = s.cell(r, 4).value
            elif t.startswith("Cumulative value") and section: blocks[section]["b"] = _num(s.cell(r, 4).value)
            elif t.startswith("Cost of material") and section: blocks[section]["c"] = _num(s.cell(r, 4).value)
            elif t.startswith("Cumulative cost") and section: blocks[section]["d"] = _num(s.cell(r, 4).value)
            elif t.startswith("V - Vna ="): vvna = _num(s.cell(r, 4).value, np.nan)
            elif t.startswith("0.96(V - Vna)"):
                f = sf.cell(r, 4).value
                mm = re.match(r"=([0-9.]+)\*", f) if isinstance(f, str) else None
                factor = float(mm.group(1)) if mm else 0.96
            elif t.startswith("Price Adjustment for"): pa = _num(s.cell(r, 4).value, np.nan)
        sigma = float((idx.px_sheet * (idx.ixc - idx.ixb) / idx.ixb).sum())
        out[int(m.group(1))] = PFSheet(int(m.group(1)), title, blocks["prev"], blocks["cur"], idx, sigma, vvna, factor, pa)
    return dict(sorted(out.items()))


# ------------------------------------------------------------------ core formulas
def px_from_result(res: Result, px_total: float = 96.0) -> pd.Series:
    """Px (%) per resource code: the resource's share of the mapped value, scaled to sum to px_total."""
    t = res.by_resource.set_index("code")["amount_lkr"]
    return t / t.sum() * px_total


def price_adjustment(px: pd.Series, ixb: pd.Series, ixc: pd.Series, v_vna: float, factor: float,
                     names: pd.Series | None = None, label: str = "", notes: list | None = None) -> PFResult:
    codes = list(ixb.index)
    p = px.reindex(codes).fillna(0.0)
    change = (ixc.reindex(codes) - ixb) / ixb
    tab = pd.DataFrame({"code": codes, "name": (names.reindex(codes).values if names is not None else codes),
                        "px": p.values, "ixb": ixb.values, "ixc": ixc.reindex(codes).values,
                        "change": change.values, "contribution": (p * change).values})
    sigma = float(tab.contribution.sum())
    return PFResult(label, p, tab, sigma, float(v_vna), factor, factor * v_vna / 100 * sigma, notes or [])


# ---------------------------------------------------------------------- calculator
class PFCalculator:
    def __init__(self, model: IPCModel, workbook: str | Path | None = None, px_total: float | None = None):
        self.model = model
        self.px_total = px_total or PX_TOTAL.get(model.profile.name, 96.0)
        self.factor = self.px_total / 100
        self.sheets = read_pf_sheets(workbook) if workbook else {}
        cum = model.targets.reindex(model.available_ipcs).cumsum()
        self._cum = cum

    # -- helpers
    def px_of_ipc(self, ipc_no: int) -> pd.Series:
        return px_from_result(self.model.resources_from_ipc(ipc_no, float(self.model.targets[ipc_no])), self.px_total)

    def valued_ipc(self, sheet: PFSheet) -> int:
        """Model IPC whose cumulative value equals the sheet's current cumulative value (fallback: N + 1)."""
        hit = self._cum.index[np.isclose(self._cum.values, sheet.cur["b"], atol=0.5)]
        return int(hit[0]) if len(hit) else sheet.number + 1

    # -- 1. reproduce the workbook sheet
    def as_sheet(self, n: int, use_sheet_factor: bool = False) -> PFResult:
        s = self.sheets[n]
        v_vna = (s.cur["b"] + s.cur["c"] - s.cur["d"]) - (s.prev["b"] + s.prev["c"] - s.prev["d"])
        res = price_adjustment(self.px_of_ipc(n), s.indices.ixb, s.indices.ixc, v_vna,
                               s.factor_sheet if use_sheet_factor else self.factor, s.indices.name,
                               f"PF {n} – as in the workbook sheet (Px of IPC {n}; V-Vna from the sheet's cumulative figures)")
        diff = np.abs(res.px.values - s.indices.px_sheet.values).max()
        res.notes.append(f"Python Px vs Px typed in the sheet: max difference {diff:.6f} %-points.")
        return res

    # -- 2. model-driven
    def model_driven(self, n: int, px_offset: int = 1) -> PFResult:
        s = self.sheets[n]
        k = self.valued_ipc(s)
        res_k = self.model.resources_from_ipc(k, float(self.model.targets[k]))
        d_mat = s.cur["c"] - s.prev["c"]
        v_vna = res_k.adjustable_base + d_mat
        px_ipc = k + (px_offset - 1)
        out = price_adjustment(self.px_of_ipc(px_ipc), s.indices.ixb, s.indices.ixc, v_vna, self.factor, s.indices.name,
                               f"PF {n} – model-driven (work of IPC {k}; Px of IPC {px_ipc}; V-Vna from the model)")
        out.notes.append(f"V - Vna = adjustable work of IPC {k} ({res_k.adjustable_base:,.2f}) + change in material at site ({d_mat:,.2f}).")
        return out

    # -- 3. any IPC value
    def from_ipc(self, ipc_no: int, value: float, ixb: pd.Series, ixc: pd.Series,
                 material_prev: float = 0.0, material_cur: float = 0.0) -> PFResult:
        res = self.model.resources_from_ipc(ipc_no, value)
        v_vna = res.adjustable_base + (material_cur - material_prev)
        names = self.model.resources.set_index("code")["name"]
        return price_adjustment(px_from_result(res, self.px_total), ixb, ixc, v_vna, self.factor, names,
                                f"IPC {ipc_no}, value {value:,.2f} – price adjustment from the model",
                                [f"V - Vna = adjustable work ({res.adjustable_base:,.2f}) + change in material at site ({material_cur - material_prev:,.2f})."])

    # -- comparison of all PF sheets
    def compare_all(self) -> pd.DataFrame:
        rows = []
        for n, s in self.sheets.items():
            a, b = self.as_sheet(n), self.model_driven(n)
            rows.append({"PF": n, "valued IPC": self.valued_ipc(s), "Excel V-Vna": s.vvna_sheet, "Excel sum": s.sigma_sheet,
                         "Excel PA": s.amount_sheet, "Python PA (as sheet)": a.amount,
                         "difference Excel-Python": s.amount_sheet - a.amount,
                         "Model V-Vna": b.v_vna, "Model sum": b.sigma, "Model-driven PA": b.amount,
                         "model-driven - sheet": b.amount - s.amount_sheet})
        return pd.DataFrame(rows)

    def non_adjustable_check(self) -> pd.DataFrame:
        """
        Cumulative non-adjustable figure (item d) of each sheet. A cumulative figure can never fall between
        valuations, and V - Vna cannot exceed the work valued plus the change in material at site.
        """
        rows = []
        for n, s in self.sheets.items():
            work, d_mat = s.cur["b"] - s.prev["b"], s.cur["c"] - s.prev["c"]
            flags = []
            if s.cur["d"] < s.prev["d"] - 0.5: flags.append("d falls")
            if s.vvna_sheet > work + d_mat + 1: flags.append("V-Vna > work + material")
            rows.append({"PF": n, "d previous": s.prev["d"], "d current": s.cur["d"], "change in d": s.cur["d"] - s.prev["d"],
                         "work valued": work, "change in material": d_mat, "V-Vna (sheet)": s.vvna_sheet, "flag": ", ".join(flags)})
        return pd.DataFrame(rows)

    def chain_check(self) -> pd.DataFrame:
        """Previous valuation of PF N must equal current valuation of PF N-1."""
        rows, ns = [], list(self.sheets)
        for a, b in zip(ns, ns[1:]):
            for key, lab in (("b", "cumulative work"), ("c", "material at site"), ("d", "non-adjustable"), ("date", "date")):
                x, y = self.sheets[a].cur.get(key), self.sheets[b].prev.get(key)
                ok = (x == y) if key == "date" else bool(np.isclose(_num(x), _num(y), atol=0.01))
                if not ok:
                    rows.append({"between": f"PF {a} current / PF {b} previous", "item": lab, "PF a current": x, "PF b previous": y})
        return pd.DataFrame(rows, columns=["between", "item", "PF a current", "PF b previous"])
