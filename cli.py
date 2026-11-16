"""
Command-line version (no web app needed).

IPC / resources (as before)
  python cli.py ipc 7 15574411.95 [--excel out.xlsx]
  python cli.py groups "Formwork=500000" "Reinforcement=1200000"
  python cli.py items 2.B.46=45000 2.E.1=300000
  python cli.py list-groups

Price fluctuation (reads the 'Price Fluctuation N' sheets of the workbook)
  python cli.py pf-all                         # Excel vs Python for every PF sheet + model-driven alternative
  python cli.py pf 3                           # PF 3 as in the workbook sheet
  python cli.py pf 3 --mode model              # PF 3 model-driven (Px of the IPC valued, V-Vna from the model)
  python cli.py pf-ipc 6 13345457.26 --indices-from 5 [--mat-prev 0 --mat-cur 500000]   # any IPC value, indices of PF sheet 5
  python cli.py pf-checks                      # chain / non-adjustable consistency checks of the PF sheets
"""
import argparse
import pandas as pd
from ipc_resource_engine import IPCModel
from pf_model import PFCalculator


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--workbook", default=None, help="workbook path (default: the corrected validation workbook)")
    common.add_argument("--excel", help="save the result to this .xlsx file")
    sub = p.add_subparsers(dest="cmd", required=True)
    a = sub.add_parser("ipc", parents=[common]); a.add_argument("ipc_no", type=int); a.add_argument("value", type=float)
    g = sub.add_parser("groups", parents=[common]); g.add_argument("pairs", nargs="+")
    i = sub.add_parser("items", parents=[common]); i.add_argument("pairs", nargs="+")
    sub.add_parser("list-groups", parents=[common])
    sub.add_parser("pf-all", parents=[common])
    sub.add_parser("pf-checks", parents=[common])
    f = sub.add_parser("pf", parents=[common]); f.add_argument("number", type=int)
    f.add_argument("--mode", choices=["sheet", "model"], default="sheet")
    f.add_argument("--px-offset", type=int, default=1, help="model mode: 1 = Px of the IPC valued (default), 0 = Px of IPC N")
    q = sub.add_parser("pf-ipc", parents=[common]); q.add_argument("ipc_no", type=int); q.add_argument("value", type=float)
    q.add_argument("--indices-from", type=int, required=True, help="PF sheet number whose base/current indices are used")
    q.add_argument("--mat-prev", type=float, default=0.0); q.add_argument("--mat-cur", type=float, default=0.0)
    args = p.parse_args()

    from ipc_resource_engine import PROFILES, DEFAULT_PROFILE
    wb = args.workbook or PROFILES[DEFAULT_PROFILE].default_workbook
    m = IPCModel.from_workbook(wb)
    pd.set_option("display.width", 250); pd.set_option("display.max_columns", 30)

    if args.cmd == "list-groups":
        print("\n".join(m.groups)); return
    if args.cmd in ("ipc", "groups", "items"):
        if args.cmd == "ipc":
            res = m.resources_from_ipc(args.ipc_no, args.value)
        else:
            d = {k.rsplit("=", 1)[0]: float(k.rsplit("=", 1)[1]) for k in args.pairs}
            res = m.resources_from_groups(d) if args.cmd == "groups" else m.resources_from_items(d)
        print(res.summary_text())
        if args.excel:
            res.to_excel(args.excel); print(f"\nSaved {args.excel}")
        return

    calc = PFCalculator(m, wb)
    if not calc.sheets:
        raise SystemExit("No 'Price Fluctuation N' sheets in this workbook.")
    if args.cmd == "pf-all":
        print(calc.compare_all().round(2).to_string(index=False))
    elif args.cmd == "pf-checks":
        print("Chain check (previous valuation of PF N = current valuation of PF N-1):")
        c = calc.chain_check(); print(c.to_string(index=False) if len(c) else "  all consistent")
        print("\nNon-adjustable / V-Vna check:"); print(calc.non_adjustable_check().round(0).to_string(index=False))
    elif args.cmd == "pf":
        r = calc.as_sheet(args.number) if args.mode == "sheet" else calc.model_driven(args.number, args.px_offset)
        print(r.summary_text())
        s = calc.sheets[args.number]
        print(f"\nWorkbook sheet says: sum {s.sigma_sheet:.4f}, V-Vna {s.vvna_sheet:,.2f}, price adjustment {s.amount_sheet:,.2f}")
        if args.excel:
            r.table.to_excel(args.excel, index=False); print(f"Saved {args.excel}")
    elif args.cmd == "pf-ipc":
        s = calc.sheets[args.indices_from]
        r = calc.from_ipc(args.ipc_no, args.value, s.indices.ixb, s.indices.ixc, args.mat_prev, args.mat_cur)
        print(r.summary_text())
        if args.excel:
            r.table.to_excel(args.excel, index=False); print(f"Saved {args.excel}")


if __name__ == "__main__":
    main()
