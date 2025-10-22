#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse, sys, re, pandas as pd, numpy as np, unicodedata
from datetime import datetime
from pathlib import Path
from collections import defaultdict

# ------------------------ logging helpers ------------------------
def now(): return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
def log(msg): print(f"[{now()}] {msg}", flush=True)

# ------------------------ column finding -------------------------
def slug(s): return re.sub(r"[\s\-_]+"," ",str(s).strip().lower()).replace(" ","")

def find_col(df, candidates):
    """Find a column in df that matches any of the candidate names (or substrings)."""
    slugs = {c: slug(c) for c in df.columns}
    inv = {v:k for k,v in slugs.items()}
    for cand in candidates:
        sc = slug(cand)
        if sc in inv: return inv[sc]
    for cand in candidates:
        sc = slug(cand)
        if len(sc)>=4:
            for c,sc2 in slugs.items():
                if sc in sc2: return c
    return None

CANDIDATES = {
    "pipcode":     ["PIPCode","PIP Code","PIP","productCode","product code"],
    "branch":      ["Branch Name","Branch","Store","Store Name"],
    "completed":   ["Completed Date","Completed","Booked Date","Booked"],
    "orderno":     ["Branch Order No.","Order No","OrderNo","Order Number"],
    "department":  ["departmentName","Department"],
    "suppliername":["supplierName","Supplier Name","Supplier Name (Display)","Supplier"],
    "groupname":   ["groupName","Group Name","Group"],
    "orderlist":   ["Supplier","Orderlist","Ordering Supplier","Supplier (Orderlist)"],
    "dns":         ["doNotStockReason","Do Not Stock Reason","DNS Reason"],
    "maxord":      ["maxOrderQuantity","Max Order Quantity","Max Ord Qty","Max Qty"],
    "req":         ["Store Order Quantity","Req Qty","Requested Qty"],
    "ord":         ["Warehouse Reply Quantity","Order Qty","Ordered Qty","Reply Qty"],
    "delv":        ["Store Received Quantity","Deliver Qty","Delivered Qty","Received Qty"],
    # product/desc info
    "desc":        ["productName","Description","Product Description","Item Description","Product Name"],
    "pack":        ["packSize","Pack Size","Pack"],
}

CURRENCY_COLUMNS = {"lineValue","cost","unitCost","NIV","Spend","Web Spend","Miscompliant Spend"}

# ------------------------ primitives -----------------------------
def ensure_numeric(s):
    out = pd.to_numeric(s, errors="coerce")
    return out.fillna(0)

def parse_date(s): return pd.to_datetime(s, errors="coerce", dayfirst=True)

def dns_present(series: pd.Series) -> pd.Series:
    s = series.astype("string")
    t = s.str.strip().str.lower()
    zero_like = t.str.fullmatch(r"0+(\.0+)?")
    placeholders = t.isin({"", "na", "n/a", "none", "null", "nan", "-", "--", "."})
    empty = t.isna() | placeholders | zero_like
    return ~empty

def autosize_sheet(writer, sheet_name, df):
    ws = writer.sheets[sheet_name]
    for idx, col in enumerate(df.columns):
        series = df[col].astype("string").fillna("")
        max_len = max([len(str(col))] + [len(x) for x in series.head(1000)])
        ws.set_column(idx, idx, min(max_len + 2, 50))

def apply_formats(writer, sheet_name, df):
    wb = writer.book
    ws = writer.sheets[sheet_name]
    pct_fmt = wb.add_format({"num_format": "0.0%"})
    cur_fmt = wb.add_format({"num_format": u"£#,##0.00"})
    for idx, col in enumerate(df.columns):
        if str(col).lower().endswith("_pct"):
            ws.set_column(idx, idx, None, pct_fmt)
    for idx, col in enumerate(df.columns):
        if str(col) in CURRENCY_COLUMNS:
            ws.set_column(idx, idx, None, cur_fmt)

def add_share_and_rate(df, qty_col, denom_col):
    if df.empty:
        df[qty_col + "_pct"] = []
        df["shortage_rate_pct"] = []
        return df
    total_qty = float(df[qty_col].sum())
    df[qty_col + "_pct"] = (df[qty_col] / total_qty) if total_qty else 0.0
    denom = df[denom_col].replace([np.inf, np.nan], 0.0)
    denom = denom.where(denom > 0, 0.0)
    df["shortage_rate_pct"] = np.where(denom > 0, df[qty_col] / denom, 0.0)
    return df

def coalesce(a, b):
    a = a.astype("string"); b = b.astype("string")
    pick_a = a.str.strip().ne("") & a.notna()
    return a.where(pick_a, b)

# robust normalizers for substitutions
def _norm_text(series: pd.Series) -> pd.Series:
    s = series.astype("string").fillna("")
    s = s.map(lambda x: unicodedata.normalize("NFKC", x))
    s = s.str.strip().str.replace(r"\s+", " ", regex=True).str.lower()
    return s

def _norm_pack(series: pd.Series) -> pd.Series:
    raw = series.astype("string")
    num = pd.to_numeric(raw.str.replace(",", "").str.extract(r"^\s*([0-9]+(?:\.[0-9]+)?)\s*$", expand=False), errors="coerce")
    num_str = num.map(lambda v: (f"{int(v)}" if pd.notna(v) and abs(v - int(v)) < 1e-9 else (f"{v:.3f}".rstrip("0").rstrip(".") if pd.notna(v) else np.nan)))
    out = num_str.astype("string")
    mask_non = out.isna() | (out.str.strip() == "nan")
    out = out.where(~mask_non, _norm_text(raw)).fillna("").str.strip()
    return out

def _make_key(desc_ser: pd.Series, pack_ser: pd.Series) -> pd.Series:
    return _norm_text(desc_ser) + " | " + _norm_pack(pack_ser)

# ------------------------ main -----------------------------
def main():
    ap = argparse.ArgumentParser(
        description="TRUE shortage report with Warehouse-only summaries, ALL-routes orderlist view, substitution exclusion, reset-on-success, NC limited to specified orderlists, earliest-date _wcDDMMYY naming."
    )
    ap.add_argument("--orders", required=True, help="Orders CSV/XLSX")
    ap.add_argument("--product-list", required=True, help="Product list CSV/XLSX")
    ap.add_argument("--out", required=True, help="Output Excel path (filename will be suffixed with wcDDMMYY)")
    ap.add_argument("--subs", help="Optional substitutions CSV/XLSX")

    # optional overrides for substitutions file columns
    ap.add_argument("--subs-desc-col", help="Column in substitutions file for original product description")
    ap.add_argument("--subs-pack-col", help="Column in substitutions file for original pack size")

    # toggles
    ap.add_argument("--completed-only", action="store_true",
                    help="If set, exclude non-completed rows (default includes them with OrderQty-as-delivery).")
    ap.add_argument("--dns-source", choices=["orders","product","both"], default="product",
                    help="Where to read doNotStockReason from.")
    ap.add_argument("--no-reset-on-success", action="store_true", help="Use monotonic Rule 3 (no reset after success).")

    # warehouse orderlists (future-proof)
    ap.add_argument(
        "--warehouse-orderlists",
        default="Warehouse;Warehouse Controlled Drugs;Warehouse - CD Products",
        help="Semicolon-separated list of orderlist names treated as Warehouse (case-insensitive)."
    )

    # NC orderlists filter (your new requirement)
    ap.add_argument(
        "--nc-orderlists",
        default="Supplier;Testers Perfume;Warehouse;Warehouse - CD Products;Xmas Warehouse;Perfumes;AAH (H&B);PHOENIX;YANKEE",
        help="Semicolon-separated orderlist names to include in Branch_NC/Company_NC (case-insensitive)."
    )

    args = ap.parse_args()
    out_path = Path(args.out); out_path.parent.mkdir(parents=True, exist_ok=True)
    errlog = out_path.parent / f"run_error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

    try:
        # ---- read inputs ----
        log(f"Reading orders: {args.orders}")
        orders = pd.read_csv(args.orders) if str(args.orders).lower().endswith(".csv") else pd.read_excel(args.orders)
        log(f"Orders rows: {len(orders)} cols: {len(orders.columns)}")

        log(f"Reading product list: {args.product_list}")
        prod = pd.read_csv(args.product_list) if str(args.product_list).lower().endswith(".csv") else pd.read_excel(args.product_list)
        log(f"Product rows: {len(prod)} cols: {len(prod.columns)}")

        subs_df = None
        if args.subs:
            log(f"Reading substitutions: {args.subs}")
            subs_df = pd.read_csv(args.subs) if str(args.subs).lower().endswith(".csv") else pd.read_excel(args.subs)
            log(f"Substitutions rows: {len(subs_df)} cols: {len(subs_df.columns)}")

        # ---- map/order columns ----
        oc = {}
        for k in ["pipcode","branch","completed","orderno","department","suppliername","groupname","orderlist","dns","maxord","req","ord","delv"]:
            oc[k] = find_col(orders, CANDIDATES[k])
            log(f"Orders mapping: {k} -> {oc[k]}")

        pc_key = find_col(prod, CANDIDATES["pipcode"])
        log(f"Product mapping: pipcode key -> {pc_key}")
        if not oc["pipcode"] or not pc_key:
            raise KeyError(f"Join key not found. Orders:{oc['pipcode']} Product:{pc_key}")

        prod_clean = prod.drop_duplicates(subset=[pc_key]).copy()
        prod_dep  = find_col(prod_clean, CANDIDATES["department"])
        prod_sup  = find_col(prod_clean, CANDIDATES["suppliername"])
        prod_grp  = find_col(prod_clean, CANDIDATES["groupname"])
        prod_ordl = find_col(prod_clean, CANDIDATES["orderlist"])
        prod_dns  = find_col(prod_clean, CANDIDATES["dns"])
        prod_maxo = find_col(prod_clean, CANDIDATES["maxord"])
        prod_desc = find_col(prod_clean, CANDIDATES["desc"])
        prod_pack = find_col(prod_clean, CANDIDATES["pack"])

        # ---- VLOOKUP merge ----
        log("Merging orders with product list...")
        merged = orders.merge(prod_clean, how="left", left_on=oc["pipcode"], right_on=pc_key, suffixes=("","__prod"))
        merged["_ProductMatched"] = merged[pc_key].notna()

        # finals for dims
        merged["Department_Final"] = coalesce(merged[oc["department"]] if oc["department"] else pd.Series("", index=merged.index, dtype="string"),
                                              merged[prod_dep] if prod_dep else pd.Series("", index=merged.index, dtype="string"))
        merged["Group_Final"]      = coalesce(merged[oc["groupname"]] if oc["groupname"] else pd.Series("", index=merged.index, dtype="string"),
                                              merged[prod_grp] if prod_grp else pd.Series("", index=merged.index, dtype="string"))

        # Orderlist: prefer Orders.Supplier (route), else Product.Orderlist
        orders_orderlist   = merged[oc["orderlist"]].astype("string") if oc["orderlist"] else pd.Series("", index=merged.index, dtype="string")
        product_orderlist  = merged[prod_ordl].astype("string") if prod_ordl else pd.Series("", index=merged.index, dtype="string")
        merged["Orderlist_Final"] = coalesce(orders_orderlist, product_orderlist)

        # SupplierName_Final = Product supplierName unless Orders suppliername is distinct from orderlist
        use_orders_supplier = oc["suppliername"] and (oc["suppliername"] != oc["orderlist"])
        orders_sup_series   = merged[oc["suppliername"]].astype("string") if use_orders_supplier else pd.Series("", index=merged.index, dtype="string")
        product_sup_series  = merged[prod_sup].astype("string") if prod_sup else pd.Series("", index=merged.index, dtype="string")
        merged["SupplierName_Final"] = coalesce(orders_sup_series, product_sup_series)

        # dates & completion
        comp = oc["completed"]
        merged["_Completed"] = parse_date(merged[comp]) if comp else pd.NaT
        merged["_IsCompleted"] = merged["_Completed"].notna()

        # numeric base; cap Req by maxOrderQuantity (if > 0)
        _req_raw = ensure_numeric(merged[oc["req"]]) if oc["req"] else pd.Series(0, index=merged.index)
        _ord     = ensure_numeric(merged[oc["ord"]]) if oc["ord"] else pd.Series(0, index=merged.index)
        _del     = ensure_numeric(merged[oc["delv"]]) if oc["delv"] else pd.Series(0, index=merged.index)
        maxo = (ensure_numeric(merged[oc["maxord"]]) if oc["maxord"] and oc["maxord"] in merged.columns
                else (ensure_numeric(merged[prod_maxo]) if prod_maxo else pd.Series(0, index=merged.index)))
        _req = np.where(maxo > 0, np.minimum(_req_raw, maxo), _req_raw)
        merged["_Req"] = pd.Series(_req, index=merged.index)
        merged["_Ord"] = _ord
        merged["_Del"] = _del

        # Effective delivery: if Completed -> Deliver; else -> Order (unless user forces completed-only)
        eff_del = np.where(merged["_IsCompleted"], _del, _ord) if not args.completed_only else _del
        merged["_EffDel"] = pd.Series(eff_del, index=merged.index)
        merged["_Short"]  = (merged["_Req"] - merged["_EffDel"]).clip(lower=0)

        # DNS consolidated
        orders_dns  = merged[oc["dns"]].astype("string") if oc["dns"] else pd.Series("", index=merged.index, dtype="string")
        product_dns = merged[prod_dns].astype("string") if prod_dns else pd.Series("", index=merged.index, dtype="string")
        if args.dns_source == "orders":
            merged["doNotStockReason_Final"] = orders_dns
        elif args.dns_source == "both":
            merged["doNotStockReason_Final"] = orders_dns.where(orders_dns.str.strip().ne(""), product_dns)
        else:
            merged["doNotStockReason_Final"] = product_dns

        # ---------------- substitutions: by ProductName+PackSize (robust) ----------------
        is_sub = pd.Series(False, index=merged.index)
        if args.subs and subs_df is not None and not subs_df.empty:
            subs_desc_col = args.subs_desc_col or find_col(subs_df, CANDIDATES["desc"]) or subs_df.columns[0]
            subs_pack_col = args.subs_pack_col or find_col(subs_df, CANDIDATES["pack"])
            if not subs_pack_col:
                subs_df["_blank_pack"] = ""
                subs_pack_col = "_blank_pack"
            subs_keys = _make_key(subs_df[subs_desc_col], subs_df[subs_pack_col])
            subs_key_set = set(subs_keys.unique())

            desc_prod_col = prod_desc if (prod_desc and prod_desc in merged.columns) else None
            pack_prod_col = prod_pack if (prod_pack and prod_pack in merged.columns) else None
            if desc_prod_col and pack_prod_col:
                orders_keys = _make_key(merged[desc_prod_col], merged[pack_prod_col])
            else:
                ord_desc_col = find_col(merged, CANDIDATES["desc"]) or "_blank_desc"
                if ord_desc_col == "_blank_desc": merged["_blank_desc"] = ""
                ord_pack_col = find_col(merged, CANDIDATES["pack"]) or "_blank_pack2"
                if ord_pack_col == "_blank_pack2": merged["_blank_pack2"] = ""
                orders_keys = _make_key(merged[ord_desc_col], merged[ord_pack_col])
            # Heuristic: substituted lines have Order Qty == 0 in the warehouse system
            is_sub = orders_keys.isin(subs_key_set) & (merged[oc["ord"]] == 0)
        merged["_IsSubstituted"] = is_sub

        # ---------------- masks & sorting ----------------
        dns_mask = dns_present(merged["doNotStockReason_Final"])          # True means has DNS (exclude)
        metric_mask = merged["_IsCompleted"] if args.completed_only else pd.Series(True, index=merged.index)

        br, pip, orderno = oc["branch"], oc["pipcode"], oc["orderno"]
        if not br or not pip: raise KeyError("Missing Branch or PIPCode for Rule 3.")
        sort_cols = [br, pip]
        if orderno: sort_cols.append(orderno)
        sort_cols.append("_Completed")
        df = merged.sort_values(by=sort_cols).copy()

        # Warehouse mask (configurable)
        allowed_wh = {x.strip().lower() for x in args.warehouse_orderlists.split(";") if x.strip()}
        ord_lower = df["Orderlist_Final"].astype("string").str.strip().str.lower()
        wh_mask = ord_lower.isin(allowed_wh)

        # reindex masks after sorting
        dns_sorted     = pd.Series(dns_mask,           index=merged.index).reindex(df.index, fill_value=False)
        metric_sorted  = pd.Series(metric_mask,        index=merged.index).reindex(df.index, fill_value=False)
        wh_sorted      = pd.Series(wh_mask,            index=df.index)
        prodmatch_sort = pd.Series(merged["_ProductMatched"].values, index=merged.index).reindex(df.index, fill_value=False)
        sub_sorted     = pd.Series(merged["_IsSubstituted"].values,  index=merged.index).reindex(df.index, fill_value=False)

        base_short = df["_Short"].astype(float).to_numpy()
        req_qty    = df["_Req"].astype(float).to_numpy()
        eff_del    = df["_EffDel"].astype(float).to_numpy()

        # candidate shortage after Rule1 + substitution (and completion toggle)
        cand = base_short.copy()
        cand = np.where(dns_sorted.to_numpy(), 0.0, cand)
        cand = np.where(sub_sorted.to_numpy(), 0.0, cand)
        if args.completed_only:
            cand = np.where(metric_sorted.to_numpy(), cand, 0.0)

        reset_on_success = not args.no_reset_on_success

        # dual tracks: WH-only (for dept/supplier/group) and ALL routes (for orderlist)
        inc_wh  = np.zeros(len(df), dtype=float)
        inc_all = np.zeros(len(df), dtype=float)
        hist_max_wh  = {}
        hist_max_all = {}

        # denominators (dedup req per window) – WH and ALL
        dedup_req_attributed = np.zeros(len(df), dtype=float)      # WH per-row attribution
        dedup_req_attributed_all = np.zeros(len(df), dtype=float)  # ALL routes per-row attribution

        window_id = np.zeros(len(df), dtype=int)
        window_seq = np.zeros(len(df), dtype=int)
        window_max_req_echo = np.zeros(len(df), dtype=float)
        denom_key_dep = np.empty(len(df), dtype=object)
        denom_key_sup = np.empty(len(df), dtype=object)
        denom_key_grp = np.empty(len(df), dtype=object)
        denom_key_ord = np.empty(len(df), dtype=object)              # WH attribution
        denom_key_orderlist_all = np.empty(len(df), dtype=object)    # ALL routes

        state_wh  = {}
        state_all = {}

        def push_window_wh(key, st, final_row_idx=None):
            """Close WH window; attribute dedup req to a row + add to WH denominators if product matched."""
            if not st: return st
            w = st["window_max_req"]
            if w > 0 and st.get("last_prod_match", False):
                idx_attr = final_row_idx if final_row_idx is not None else st.get("last_row_idx")
                if idx_attr is not None:
                    dedup_req_attributed[idx_attr] = w
                    window_max_req_echo[idx_attr] = w
                    denom_key_dep[idx_attr] = st["last_dep"]
                    denom_key_sup[idx_attr] = st["last_sup"]
                    denom_key_grp[idx_attr] = st["last_grp"]
                    denom_key_ord[idx_attr] = st["last_ord"]
            st["window_max_req"] = 0.0
            st["open"] = False
            return st

        def push_window_all(key, st, final_row_idx=None):
            """Close ALL-routes window; add to ALL orderlist denominator."""
            if not st: return st
            w = st["window_max_req"]
            if w > 0 and st["last_ord"] is not None:
                idx_attr = final_row_idx if final_row_idx is not None else st.get("last_row_idx")
                if idx_attr is not None:
                    denom_key_orderlist_all[idx_attr] = st["last_ord"]
                    dedup_req_attributed_all[idx_attr] = w
            st["window_max_req"] = 0.0
            st["open"] = False
            return st

        dep_final = df["Department_Final"].astype("string").replace("", pd.NA).to_numpy()
        sup_final = df["SupplierName_Final"].astype("string").replace("", pd.NA).to_numpy()
        grp_final = df["Group_Final"].astype("string").replace("", pd.NA).to_numpy()
        ord_final = df["Orderlist_Final"].astype("string").replace("", pd.NA).to_numpy()

        for i, (b, pp, c, sb, rq, dval, sval, gval, oval,
                is_wh, is_dns, prod_ok, is_subbed, ok_completed_rule, edel) in enumerate(
            zip(df[br].to_numpy(),
                df[pip].astype(str).to_numpy(),
                cand, base_short, req_qty, dep_final, sup_final, grp_final, ord_final,
                wh_sorted.to_numpy(), dns_sorted.to_numpy(), prodmatch_sort.to_numpy(),
                sub_sorted.to_numpy(), metric_sorted.to_numpy(), eff_del)
        ):
            key = (b, pp)

            # Substituted rows are neutral for shortages/denominators.
            if is_subbed:
                inc_wh[i] = 0.0; inc_all[i] = 0.0
                continue

            # ---------- WH track ----------
            stw = state_wh.get(key, {
                "window_max_req": 0.0, "last_dep": None, "last_sup": None, "last_grp": None, "last_ord": None,
                "open": False, "window_id": 0, "seq": 0, "last_row_idx": None, "last_prod_match": False
            })
            if is_wh:
                if dval is not pd.NA: stw["last_dep"] = dval
                if sval is not pd.NA: stw["last_sup"] = sval
                if gval is not pd.NA: stw["last_grp"] = gval
                if oval is not pd.NA: stw["last_ord"] = oval
                stw["last_prod_match"] = bool(prod_ok)

            if ok_completed_rule:
                if is_wh and not stw["open"]:
                    stw["open"] = True; stw["window_id"] += 1; stw["seq"] = 0
                if is_wh:
                    stw["seq"] += 1; stw["last_row_idx"] = i
                    window_id[i] = stw["window_id"]; window_seq[i] = stw["seq"]
                if is_wh and not is_dns:
                    stw["window_max_req"] = max(stw["window_max_req"], float(rq))

                # success (Rule-3 reset) if no short OR req <= effective delivery
                if reset_on_success and (sb <= 0 or c <= 0 or rq <= edel):
                    if is_wh: stw = push_window_wh(key, stw, final_row_idx=i)
                    hist_max_wh[key] = 0.0; inc_wh[i] = 0.0
                else:
                    prev_max = hist_max_wh.get(key, 0.0)
                    add = max(0.0, c - prev_max) if c > prev_max else 0.0
                    hist_max_wh[key] = max(hist_max_wh.get(key, 0.0), c)
                    inc_wh[i] = add if (is_wh and prod_ok) else 0.0
            else:
                inc_wh[i] = 0.0
            state_wh[key] = stw

            # ---------- ALL routes track ----------
            sta = state_all.get(key, {"window_max_req": 0.0, "last_ord": None, "open": False, "last_row_idx": None})
            if ok_completed_rule:
                if not sta["open"]:
                    sta["open"] = True; sta["window_max_req"] = 0.0
                if oval is not pd.NA: sta["last_ord"] = oval
                sta["last_row_idx"] = i
                if not is_dns:
                    sta["window_max_req"] = max(sta["window_max_req"], float(rq))

                if reset_on_success and (sb <= 0 or c <= 0 or rq <= edel):
                    sta = push_window_all(key, sta, final_row_idx=i)
                    hist_max_all[key] = 0.0; inc_all[i] = 0.0
                else:
                    prev_max_all = hist_max_all.get(key, 0.0)
                    add_all = max(0.0, c - prev_max_all) if c > prev_max_all else 0.0
                    hist_max_all[key] = max(hist_max_all.get(key, 0.0), c)
                    inc_all[i] = add_all
            else:
                inc_all[i] = 0.0
            state_all[key] = sta

        # close open windows
        for key, stw in list(state_wh.items()):
            if stw.get("open", False): push_window_wh(key, stw, final_row_idx=stw.get("last_row_idx"))
        for key, sta in list(state_all.items()):
            if sta.get("open", False): push_window_all(key, sta, final_row_idx=sta.get("last_row_idx"))

        # attach results
        df["TrueShortQty_WH"]  = inc_wh
        df["TrueShortQty_ALL"] = inc_all
        df["window_id"] = window_id
        df["window_seq_incr"] = window_seq
        df["window_max_req"] = window_max_req_echo
        df["dedup_req_attributed"] = dedup_req_attributed
        df["dedup_req_attributed_all"] = dedup_req_attributed_all
        df["denom_key_department"] = pd.Series(denom_key_dep, index=df.index)
        df["denom_key_supplier"]   = pd.Series(denom_key_sup, index=df.index)
        df["denom_key_group"]      = pd.Series(denom_key_grp, index=df.index)
        df["denom_key_orderlist"]  = pd.Series(denom_key_ord, index=df.index)               # WH attribution
        df["denom_key_orderlist_all"] = pd.Series(denom_key_orderlist_all, index=df.index)  # ALL routes
        df["_ProductMatched"] = prodmatch_sort.to_numpy()
        df["_IsSubstituted"] = sub_sorted.to_numpy()
        df["_Effective_Delivery_Used"] = df["_EffDel"]

        # roll masks
        roll_mask_wh  = (~df["_IsSubstituted"]) & (wh_sorted & prodmatch_sort)  # WH-only & matched & not subbed
        roll_mask_all = (~df["_IsSubstituted"])                                  # ALL routes & not subbed

        # ---------- Build denominators from attributed rows ----------
        def denom_from_keys(series_key, values):
            tmp = pd.DataFrame({"k": series_key, "v": values})
            out = tmp.groupby("k", dropna=False)["v"].sum().reset_index()
            out = out.rename(columns={"k":"key","v":"dedup_req_qty"})
            return dict(zip(out["key"], out["dedup_req_qty"]))

        # WH-only denominators (use dedup_req_attributed and WH keys)
        denom_dep = denom_from_keys(df["denom_key_department"], df["dedup_req_attributed"])
        denom_sup = denom_from_keys(df["denom_key_supplier"],   df["dedup_req_attributed"])
        denom_grp = denom_from_keys(df["denom_key_group"],      df["dedup_req_attributed"])

        # ALL routes denominators (use ALL orderlist per-row attributions)
        denom_ord_all = denom_from_keys(df["denom_key_orderlist_all"], df["dedup_req_attributed_all"])

        # ---------- rollups ----------
        def rollup(mask, series, title, qty_col, denom_map):
            cols = [title, "true_short_qty", "true_short_lines", "dedup_req_qty",
                    "true_short_qty_pct", "shortage_rate_pct"]
            if series is None: return pd.DataFrame(columns=cols)
            key_series = series.rename(title)
            g = (
                df.loc[mask.to_numpy()]
                  .groupby(key_series, dropna=False)
                  .agg(
                      true_short_qty=(qty_col,"sum"),
                      true_short_lines=(qty_col, lambda s: (s>0).sum())
                  )
                  .reset_index()
            )
            dd = pd.DataFrame([(k, v) for k, v in denom_map.items()],
                              columns=[title, "dedup_req_qty"])
            g = g.merge(dd, on=title, how="left")
            g["dedup_req_qty"] = g["dedup_req_qty"].fillna(0).astype(float)
            g = add_share_and_rate(g, "true_short_qty", "dedup_req_qty")
            return g.sort_values(by=["true_short_qty","true_short_lines"], ascending=False)

        # warehouse-only summaries (exclude non-matched PIPs)
        r_dep = rollup(roll_mask_wh,  df["Department_Final"],   "Department",   "TrueShortQty_WH",  denom_dep)
        r_sup = rollup(roll_mask_wh,  df["SupplierName_Final"], "SupplierName", "TrueShortQty_WH",  denom_sup)
        r_grp = rollup(roll_mask_wh,  df["Group_Final"],        "Group",        "TrueShortQty_WH",  denom_grp)

        # Orderlist (ALL routes)
        r_ord = rollup(roll_mask_all, df["Orderlist_Final"], "Orderlist", "TrueShortQty_ALL", denom_ord_all)

        # ---------- NC (line-based) restricted to selected orderlists ----------
        nc_allowed = {x.strip().lower() for x in args.nc_orderlists.split(";") if x.strip()}
        ordlist_lower = df["Orderlist_Final"].astype("string").str.strip().str.lower()
        nc_mask = ordlist_lower.isin(nc_allowed)
        df_nc = df.loc[nc_mask].copy()

        branch = (
            df_nc.groupby(oc["branch"], dropna=False)
                 .size().reset_index(name="total_lines")
        )
        non_completed = (
            df_nc.loc[~df_nc["_IsCompleted"]]
                 .groupby(oc["branch"], dropna=False)
                 .size().reset_index(name="non_completed_lines")
        )
        branch = branch.merge(non_completed, on=oc["branch"], how="left").fillna({"non_completed_lines": 0})
        branch["non_completed_lines"] = branch["non_completed_lines"].astype(int)
        branch["non_completed_pct"] = np.where(
            branch["total_lines"] > 0,
            branch["non_completed_lines"] / branch["total_lines"],
            0.0
        )

        company_total = int(branch["total_lines"].sum())
        company_nc    = int(branch["non_completed_lines"].sum())
        company = pd.DataFrame([{
            "company_total_lines": company_total,
            "company_non_completed_lines": company_nc,
            "company_non_completed_pct": (company_nc/company_total) if company_total>0 else 0.0
        }])

        # mismatch detail/summary
        _ord2 = ensure_numeric(df[oc["ord"]]); _del2 = ensure_numeric(df[oc["delv"]])
        mis_mask = (_ord2 != _del2)
        cols_keep = [c for c in [oc["pipcode"], oc["department"], oc["branch"], oc["req"], oc["ord"], oc["delv"], oc["completed"]] if c]
        mis_detail = df.loc[mis_mask, cols_keep] if cols_keep else df.loc[mis_mask]
        if oc["pipcode"] and oc["department"] and oc["branch"]:
            mis_summary = df.loc[mis_mask].groupby([oc["pipcode"], oc["department"], oc["branch"]], dropna=False).size().reset_index(name="lines")
        else:
            if oc["pipcode"]:
                mis_summary = df.loc[mis_mask].groupby([oc["pipcode"]], dropna=False).size().reset_index(name="lines")
            else:
                mis_summary = pd.DataFrame(columns=["lines"])

        # DNS Top Reasons (WH-only, matched, not substituted)
        top_dns = pd.DataFrame()
        wh_dns_mask = (roll_mask_wh.to_numpy())
        wh_dns = df.loc[wh_dns_mask, "doNotStockReason_Final"].astype("string").str.strip().str.lower()
        wh_dns = wh_dns[(wh_dns.notna()) & (wh_dns!="") & (~wh_dns.isin({"0","na","n/a","none","null","nan","-","--","."}))]
        if not wh_dns.empty:
            top_dns = wh_dns.value_counts().reset_index()
            top_dns.columns = ["doNotStockReason","lines"]

        # Diagnostics
        diag = {
            "rows_total": int(len(df)),
            "rows_completed": int(df["_IsCompleted"].sum()),
            "rows_non_completed": int((~df["_IsCompleted"]).sum()),
            "warehouse_rows": int(wh_sorted.sum()),
            "warehouse_rows_matched": int((wh_sorted & prodmatch_sort).sum()),
            "substituted_rows": int(df["_IsSubstituted"].sum()),
            "base_short_lines": int((df["_Short"]>0).sum()),
            "base_short_qty": float(df.loc[df["_Short"]>0, "_Short"].sum()),
            "final_true_short_lines_WH": int((df["TrueShortQty_WH"]>0).sum()),
            "final_true_short_qty_WH": float(df["TrueShortQty_WH"].sum()),
            "final_true_short_lines_ALL": int((df["TrueShortQty_ALL"]>0).sum()),
            "final_true_short_qty_ALL": float(df["TrueShortQty_ALL"].sum()),
        }
        diag_df = pd.DataFrame([{"metric": k, "value": v} for k,v in diag.items()])

        # Top Short Lines (WH-only, matched, not substituted)
        mask_top = (roll_mask_wh.to_numpy()) & (df["TrueShortQty_WH"] > 0)
        cols = [
            oc["branch"], oc["completed"], oc["orderno"], oc["orderlist"],
            oc["suppliername"], oc["pipcode"], oc["req"], oc["ord"], oc["delv"],
            "Department_Final", "SupplierName_Final", "Group_Final", "Orderlist_Final",
            "TrueShortQty_WH","_Short","_Req","_Ord","_Del","_EffDel",
        ]
        if oc["department"]: cols.insert(4, oc["department"])
        if oc["groupname"]:  cols.insert(6, oc["groupname"])
        cols = [c for c in cols if c]
        top = df.loc[mask_top, cols].copy()
        rename_map = {}
        if oc["branch"]:        rename_map[oc["branch"]]        = "Branch"
        if oc["completed"]:     rename_map[oc["completed"]]     = "Completed Date"
        if oc["orderno"]:       rename_map[oc["orderno"]]       = "Branch Order No."
        if oc["orderlist"]:     rename_map[oc["orderlist"]]     = "Orderlist (Raw)"
        if oc["department"]:    rename_map[oc["department"]]    = "Department (Raw)"
        if oc["suppliername"]:  rename_map[oc["suppliername"]]  = "SupplierName (Raw)"
        if oc["groupname"]:     rename_map[oc["groupname"]]     = "Group (Raw)"
        if oc["pipcode"]:       rename_map[oc["pipcode"]]       = "PIPCode"
        if oc["req"] :          rename_map[oc["req"]]           = "Req Qty (capped)"
        if oc["ord"] :          rename_map[oc["ord"]]           = "Order Qty"
        if oc["delv"]:          rename_map[oc["delv"]]          = "Deliver Qty"
        top = top.rename(columns=rename_map).sort_values(by="TrueShortQty_WH", ascending=False)

        # Top Short PIPs (WH-only, matched, not substituted) with description/pack for ID
        pip_col = oc["pipcode"]
        if pip_col:
            df_top = df.loc[mask_top]
            desc_col = prod_desc if (prod_desc and prod_desc in df_top.columns) else find_col(df_top, CANDIDATES["desc"])
            pack_col = prod_pack if (prod_pack and prod_pack in df_top.columns) else find_col(df_top, CANDIDATES["pack"])
            agg_dict = {
                "true_short_qty": ("TrueShortQty_WH","sum"),
                "true_short_lines": ("TrueShortQty_WH", lambda s: (s>0).sum()),
            }
            if oc["branch"]:
                agg_dict["branches_involved"] = (oc["branch"], pd.Series.nunique)
            if desc_col:
                agg_dict["Product_Description"] = (desc_col, lambda s: s.dropna().astype(str).iloc[0] if s.dropna().size else "")
            if pack_col:
                agg_dict["Pack_Size"] = (pack_col, lambda s: s.dropna().astype(str).iloc[0] if s.dropna().size else "")
            pip_agg = (df_top.groupby(pip_col, dropna=False)
                             .agg(**agg_dict)
                             .reset_index())
            pip_agg = pip_agg.sort_values("true_short_qty", ascending=False)
            total_ts_pip = float(pip_agg["true_short_qty"].sum())
            pip_agg["true_short_qty_pct"] = np.where(total_ts_pip>0, pip_agg["true_short_qty"]/total_ts_pip, 0.0)
        else:
            pip_agg = pd.DataFrame(columns=["PIPCode","true_short_qty","true_short_lines","branches_involved",
                                            "true_short_qty_pct","Product_Description","Pack_Size"])

        # ---- final filename wcDDMMYY from earliest Completed Date ----
        earliest = pd.to_datetime(df["_Completed"], errors="coerce").dropna()
        if not earliest.empty:
            d0 = earliest.min()
            wc_tag = f"wc{d0.day:02d}{d0.month:02d}{d0.year%100:02d}"
            base = out_path.with_suffix("").name
            parent = out_path.parent
            if not base.lower().endswith(f"_{wc_tag}"):
                final_name = f"{base}_{wc_tag}.xlsx"
            else:
                final_name = f"{base}.xlsx"
            out_final = parent / final_name
        else:
            out_final = out_path

        # ---- write workbook ----
        log(f"Writing Excel: {out_final}")
        with pd.ExcelWriter(out_final, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as xw:
            sheets = [
                ("Dept_Shortage", r_dep),                # WH only, matched, no subs
                ("Supplier_Shortage", r_sup),            # WH only, matched, no subs
                ("Group_Shortage", r_grp),               # WH only, matched, no subs
                ("Orderlist_Shortage", r_ord),           # ALL routes, no subs
                ("Top_Short_PIPs", pip_agg),             # WH only, matched
                ("Top_Short_Lines", top),                # WH only, matched
                ("Mismatch_Detail", mis_detail),
                ("Mismatch_Summary", mis_summary),
                ("Branch_NC", branch),
                ("Company_NC", company),
                ("DNS_Top_Reasons", top_dns),
                ("Orders_Enriched", df),
                ("Diagnostics", pd.DataFrame([{"metric": k, "value": v} for k,v in diag.items()])),
            ]
            for name, data in sheets:
                data.to_excel(xw, index=False, sheet_name=name)
                autosize_sheet(xw, name, data)
                apply_formats(xw, name, data)

        log("DONE.")
        return 0

    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        with open(errlog, "w", encoding="utf-8") as fh:
            fh.write(tb)
        print(tb, file=sys.stderr)
        print(f"[!] Error logged to: {errlog}", file=sys.stderr)
        return 2

if __name__=="__main__":
    sys.exit(main())
