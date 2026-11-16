"""
IPC -> Resource Composition engine  (version 2: works for any project built on the framework)
===========================================================================================
Reads the framework workbook once and answers, for any IPC value or any completed work,
"which resources were consumed and how much?".

Project profiles
----------------
The framework is the same in every project; only sheet/column names differ. A profile
(ProjectProfile) tells the engine where to find things. Two profiles are built in:
  * "validation" – Validation_-_IPC_Dynamic_Allocation_and_Framework(_CORRECTED).xlsx
  * "project1"   – ori_5__Allocation_of_IPC_s_and_the_IPC_related_Compositions(_CORRECTED).xlsx
The profile is detected automatically from the sheet names.

Data used
---------
* Resource composition % of every BOQ work group  ('Resource Composition Percentage')
* BOQ rows: group, treatment (ADJUSTABLE / NON-ADJUSTABLE / DISREGARDED) and the IPC phasing
  profile ('17 IPC BOQ Workdone')
* IPC period values (targets) of the IPCs in the workbook

Three ways to use it
--------------------
1. resources_from_ipc(ipc_no, value)          only the IPC value is known
2. resources_from_groups({"Formwork": 5e5})   work completed per work group
3. resources_from_rows({row_id: amount}) / resources_from_items({"2.B.46": amount})   per BOQ row / item

Plus price_fluctuation(result, base_index, current_index) and merge_results([...]).
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd

ADJ, NON_ADJ, DISREG = "ADJUSTABLE", "NON-ADJUSTABLE", "DISREGARDED"


# ------------------------------------------------------------------ profiles
@dataclass(frozen=True)
class ProjectProfile:
    name: str
    default_workbook: str
    detect_sheet: str                 # a sheet that identifies this project type
    comp_sheet: str = "Resource Composition Percentage"
    comp_group_col: str = "Group"
    comp_first_resource: str = "L1 – Skilled Labour"
    comp_last_resource: str = "P3 – Fuel"
    comp_scale: float = 1.0           # divide composition cells by this (100 if stored as percent points)
    items_sheet: str = "17 IPC BOQ Workdone"
    id_col: str = "Record ID"
    item_col: str = "BOQ Item"
    desc_col: str = "Description"
    group_col: str = "BOQ Group"
    treat_col: str = "Final Treatment"
    amount_col: str = "BOQ Amount LKR"
    ipc_col_regex: str = r"IPC (\d+) LKR"
    targets_sheet: str = "IPC Targets"
    target_ipc_col: str = "IPC No"
    target_value_col: str = "Analytical IPC Workdone Target LKR"


PROFILES = {
    "project1": ProjectProfile(
        name="project1",
        default_workbook="ori_5__Allocation_of_IPC_s_and_the_IPC_related_Compositions_CORRECTED.xlsx",
        detect_sheet="IPC Targets"),
    "validation": ProjectProfile(
        name="validation",
        default_workbook="VALIDATION_of_the_Developed_Model_CORRECTED.xlsx",
        detect_sheet="IPC Verified Targets",
        comp_last_resource="M18 – Sanitary Ware (Local)", comp_scale=100.0,
        id_col="ID", treat_col="Treatment",
        targets_sheet="IPC Verified Targets", target_value_col="Period BOQ V"),
}
DEFAULT_PROFILE = "validation"
DEFAULT_WORKBOOK = PROFILES[DEFAULT_PROFILE].default_workbook


def detect_profile(path) -> ProjectProfile:
    names = pd.ExcelFile(path).sheet_names
    for p in PROFILES.values():
        if p.detect_sheet in names:
            return p
    raise ValueError("Workbook not recognised – add a ProjectProfile for it.")


def _clean(s) -> str:
    """Normalise group names (stray spaces are common in the workbooks)."""
    return re.sub(r"\s+", " ", str(s)).strip()


def _resource_type(code: str) -> str:
    return {"L": "Labour", "M": "Material", "P": "Plant"}.get(code[0], "Other")


# --------------------------------------------------------------------- model
@dataclass
class IPCModel:
    profile: ProjectProfile
    comp: pd.DataFrame              # index = group, columns = resource labels, values = fraction (rows sum to ~1)
    items: pd.DataFrame
    ipc_cols: list
    targets: pd.Series              # IPC no -> period value (LKR)
    resources: pd.DataFrame = field(init=False)

    def __post_init__(self):
        rows = []
        for lab in self.comp.columns:
            code, _, name = lab.partition(" – ")
            rows.append({"label": lab, "code": code.strip(), "name": name.strip(), "type": _resource_type(code.strip())})
        self.resources = pd.DataFrame(rows).set_index("label")

    @classmethod
    def from_workbook(cls, path: str | Path | None = None, profile: str | None = None) -> "IPCModel":
        import io
        if path is None:
            p = PROFILES[profile or DEFAULT_PROFILE]
            path = p.default_workbook
        if hasattr(path, "read"):                       # uploaded file (Streamlit / BytesIO)
            data = path.read()
            src = lambda: io.BytesIO(data)
        else:
            src = lambda: str(path)
        p = PROFILES[profile] if profile else detect_profile(src())

        comp = pd.read_excel(src(), sheet_name=p.comp_sheet)
        comp[p.comp_group_col] = comp[p.comp_group_col].map(_clean)
        cols = list(comp.columns)
        res_cols = cols[cols.index(p.comp_first_resource): cols.index(p.comp_last_resource) + 1]
        comp = comp.set_index(p.comp_group_col)[res_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0) / p.comp_scale
        comp = comp[~comp.index.duplicated(keep="first")]

        items = pd.read_excel(src(), sheet_name=p.items_sheet)
        items[p.group_col] = items[p.group_col].map(_clean)
        items[p.item_col] = items[p.item_col].astype(str).str.strip()
        items[p.amount_col] = pd.to_numeric(items[p.amount_col], errors="coerce").fillna(0.0)   # 'Rate Only' -> 0
        ipc_cols = [c for c in items.columns if re.fullmatch(p.ipc_col_regex, str(c))]
        items[ipc_cols] = items[ipc_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)

        tg = pd.read_excel(src(), sheet_name=p.targets_sheet)
        nos = pd.to_numeric(tg[p.target_ipc_col], errors="coerce")
        vals = pd.to_numeric(tg[p.target_value_col], errors="coerce")
        ok = nos.notna() & vals.notna()
        targets = pd.Series(vals[ok].values, index=nos[ok].astype(int).values)
        return cls(profile=p, comp=comp, items=items, ipc_cols=ipc_cols, targets=targets)

    # ------------------------------------------------------------- utilities
    @property
    def groups(self) -> list:
        return list(self.comp.index)

    @property
    def available_ipcs(self) -> list:
        return [int(re.fullmatch(self.profile.ipc_col_regex, c).group(1)) for c in self.ipc_cols]

    def composition_of(self, group: str) -> pd.Series:
        return self.comp.loc[_clean(group)]

    def _group_resource_matrix(self, group_amounts: pd.Series) -> pd.DataFrame:
        g = group_amounts[group_amounts.index.isin(self.comp.index)]
        return self.comp.loc[g.index].mul(g, axis=0)

    def _package(self, group_amounts: pd.Series, input_value: float, notes: list,
                 non_adjustable=0.0, disregarded=0.0, unmapped=0.0) -> "Result":
        matrix = self._group_resource_matrix(group_amounts)
        table = self.resources.copy()
        table["amount_lkr"] = matrix.sum()
        base = float(group_amounts.sum())
        table["share_of_adjustable_pct"] = (table["amount_lkr"] / base * 100) if base else 0.0
        return Result(input_value, base, non_adjustable, disregarded, unmapped, table.reset_index(drop=True),
                      matrix.loc[:, (matrix != 0).any()], group_amounts, notes)

    # ------------------------------------------------ 1. entered by GROUP
    def resources_from_groups(self, group_amounts: dict) -> "Result":
        s = pd.Series({_clean(k): float(v) for k, v in group_amounts.items()}, dtype=float)
        unknown = [g for g in s.index if g not in self.comp.index]
        if unknown:
            raise KeyError(f"Unknown work group(s): {unknown}. See model.groups for valid names.")
        return self._package(s, float(s.sum()), ["Mode: work-done entered per BOQ work group."])

    # ------------------------------------------- 2. entered by ROW or ITEM
    def _from_row_amounts(self, row_amounts: pd.Series, notes: list) -> "Result":
        p, it = self.profile, self.items
        group_tot, na, dis, unm = {}, 0.0, 0.0, 0.0
        for idx, v in row_amounts.items():
            row = it.loc[idx]
            t = row[p.treat_col]
            if t == NON_ADJ:
                na += v
            elif t == DISREG:
                dis += v
            elif t == ADJ and row[p.group_col] in self.comp.index:
                group_tot[row[p.group_col]] = group_tot.get(row[p.group_col], 0.0) + v
            else:
                unm += v
                notes.append(f"Row {row[p.id_col]} (item {row[p.item_col]}): '{row[p.group_col]}' has no composition – not mapped.")
        return self._package(pd.Series(group_tot, dtype=float), float(row_amounts.sum()), notes, na, dis, unm)

    def resources_from_rows(self, row_amounts: dict) -> "Result":
        """row_amounts: {unique row ID (sheet 17 ID column): LKR completed}."""
        p, it = self.profile, self.items
        by_id = it.reset_index().set_index(p.id_col)["index"]
        s = {}
        for k, v in row_amounts.items():
            if k not in by_id.index:
                raise KeyError(f"Row ID '{k}' not found.")
            s[by_id[k]] = float(v)
        return self._from_row_amounts(pd.Series(s, dtype=float), ["Mode: work-done entered per BOQ row."])

    def resources_from_items(self, item_amounts: dict) -> "Result":
        """
        item_amounts: {BOQ item no.: LKR completed}.
        Items that were split into components (e.g. lintels 2.B.46 -> concrete, formwork, steel, plaster)
        have several rows; the amount is spread over them in proportion to their BOQ amounts.
        """
        p, it = self.profile, self.items
        s = {}
        for k, v in item_amounts.items():
            rows = it[it[p.item_col] == str(k).strip()]
            if rows.empty:
                raise KeyError(f"BOQ item '{k}' not found.")
            w = rows[p.amount_col]
            w = w / w.sum() if w.sum() > 0 else pd.Series(1.0 / len(rows), index=rows.index)
            for idx, share in w.items():
                s[idx] = s.get(idx, 0.0) + float(v) * float(share)
        return self._from_row_amounts(pd.Series(s, dtype=float), ["Mode: work-done entered per BOQ item."])

    # --------------------------------------------- 3. entered as IPC VALUE
    def ipc_profile(self, ipc_no: int | None) -> pd.DataFrame:
        """
        Item-level phasing profile (weights sum to 1).
        * IPC in the workbook -> that IPC's allocation in sheet 17
        * otherwise           -> uniform pro-rata to BOQ amounts
        """
        p, it = self.profile, self.items
        if ipc_no in self.available_ipcs:
            raw = it[f"IPC {ipc_no} LKR"]
            note = f"profile of IPC {ipc_no} from sheet 17"
        else:
            raw = it[p.amount_col].where(it[p.treat_col].isin([ADJ, NON_ADJ, DISREG]), 0.0)
            note = "uniform BOQ pro-rata profile (IPC not in the workbook)"
        prof = pd.DataFrame({"group": it[p.group_col], "treatment": it[p.treat_col], "raw": raw})
        prof["weight"] = prof["raw"] / prof["raw"].sum()
        prof.attrs["note"] = note
        return prof

    def resources_from_ipc(self, ipc_no: int, value: float) -> "Result":
        """value = IPC *period* workdone value (LKR), not cumulative."""
        prof = self.ipc_profile(ipc_no)
        prof["lkr"] = prof["weight"] * float(value)
        na = float(prof.loc[prof.treatment == NON_ADJ, "lkr"].sum())
        dis = float(prof.loc[prof.treatment == DISREG, "lkr"].sum())
        adj = prof[prof.treatment == ADJ]
        mapped = adj[adj.group.isin(self.comp.index)]
        unm = float(adj["lkr"].sum() - mapped["lkr"].sum())
        return self._package(mapped.groupby("group")["lkr"].sum(), float(value),
                             [f"Mode: IPC value only – {prof.attrs['note']}."], na, dis, unm)


# ------------------------------------------------------------------- result
@dataclass
class Result:
    input_value: float
    adjustable_base: float
    non_adjustable: float
    disregarded: float
    unmapped: float
    by_resource: pd.DataFrame
    by_group: pd.DataFrame
    group_amounts: pd.Series
    notes: list

    @property
    def total_resources(self) -> float:
        return float(self.by_resource["amount_lkr"].sum())

    def nonzero(self) -> pd.DataFrame:
        t = self.by_resource
        return t[t["amount_lkr"].abs() > 1e-9].sort_values("amount_lkr", ascending=False).reset_index(drop=True)

    def by_type(self) -> pd.Series:
        return self.by_resource.groupby("type")["amount_lkr"].sum()

    def summary_text(self) -> str:
        lines = [
            f"Input value                    : {self.input_value:>18,.2f}",
            f"  Non-adjustable portion       : {self.non_adjustable:>18,.2f}",
            f"  Disregarded portion          : {self.disregarded:>18,.2f}",
            f"  Adjustable base              : {self.adjustable_base:>18,.2f}",
            f"Total resources mapped         : {self.total_resources:>18,.2f}",
            "",
            f"{'Code':<6}{'Resource':<42}{'Type':<10}{'LKR':>16}{'% of base':>11}",
        ]
        for _, r in self.nonzero().iterrows():
            lines.append(f"{r.code:<6}{r['name'][:40]:<42}{r.type:<10}{r.amount_lkr:>16,.2f}{r.share_of_adjustable_pct:>10.2f}%")
        return "\n".join(lines + [""] + [f"Note: {n}" for n in self.notes])

    def to_excel(self, path):
        with pd.ExcelWriter(path, engine="openpyxl") as xw:
            pd.DataFrame({
                "Item": ["Input value", "Non-adjustable", "Disregarded", "Adjustable base", "Total resources mapped", "Unmapped adjustable"],
                "LKR": [self.input_value, self.non_adjustable, self.disregarded, self.adjustable_base,
                        self.total_resources, self.unmapped]}).to_excel(xw, sheet_name="Summary", index=False)
            self.nonzero().to_excel(xw, sheet_name="By resource", index=False)
            self.by_group.to_excel(xw, sheet_name="By group x resource")
            self.group_amounts.rename("Adjustable LKR").to_frame().to_excel(xw, sheet_name="Work groups")


# --------------------------------------------------------- price fluctuation
def price_fluctuation(result: Result, base_index: dict, current_index: dict) -> pd.DataFrame:
    """fluctuation_i = amount_i * (current_i / base_i - 1); indices keyed by resource code ("L1", "M3" …)."""
    t = result.nonzero().copy()
    t["base_index"] = t["code"].map(base_index)
    t["current_index"] = t["code"].map(current_index)
    # Streamlit's editor (and callers outside the UI) can supply numeric-looking
    # values as strings. Normalize them before arithmetic; missing or invalid
    # values leave the resource unchanged. A zero base index is undefined too.
    base = pd.to_numeric(t["base_index"], errors="coerce").to_numpy(dtype="float64", na_value=np.nan)
    current = pd.to_numeric(t["current_index"], errors="coerce").to_numpy(dtype="float64", na_value=np.nan)
    valid = np.isfinite(base) & np.isfinite(current) & (base != 0)
    ratio = np.divide(current, base, out=np.ones_like(current), where=valid)
    t["change_pct"] = (ratio - 1) * 100
    amount = pd.to_numeric(t["amount_lkr"], errors="coerce").to_numpy(dtype="float64", na_value=0.0)
    t["fluctuation_lkr"] = amount * (ratio - 1)
    return t


def merge_results(results: list) -> Result:
    """Add several Results together (work completed in several sessions)."""
    base = results[0]
    table = base.by_resource.copy()
    table["amount_lkr"] = sum(r.by_resource["amount_lkr"].values for r in results)
    adj = sum(r.adjustable_base for r in results)
    table["share_of_adjustable_pct"] = table["amount_lkr"] / adj * 100 if adj else 0.0
    by_group = results[0].by_group
    for r in results[1:]:
        by_group = by_group.add(r.by_group, fill_value=0.0)
    ga = results[0].group_amounts
    for r in results[1:]:
        ga = ga.add(r.group_amounts, fill_value=0.0)
    return Result(sum(r.input_value for r in results), adj, sum(r.non_adjustable for r in results),
                  sum(r.disregarded for r in results), sum(r.unmapped for r in results), table, by_group, ga,
                  [n for r in results for n in r.notes])


if __name__ == "__main__":
    import sys
    m = IPCModel.from_workbook(sys.argv[1] if len(sys.argv) > 1 else None)
    k = m.available_ipcs[min(5, len(m.available_ipcs) - 1)]
    print(m.resources_from_ipc(k, float(m.targets[k])).summary_text())
