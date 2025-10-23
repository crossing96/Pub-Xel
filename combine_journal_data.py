import pandas as pd
import numpy as np
import re

# ---------- CONFIG ----------
FILE_FIRST = "data/journal_abb.txt"
FILE_SECOND = "data/journal_pub.txt"
FILE_THIRD = "data/journal_if.txt"
FILE_OUT = "data/journal_combined.txt"
JOURNAL_PUB_SOURCE = "Journal Name"   # or "Abbreviated Journal"
# ----------------------------

def normalize_issn(x: str) -> str | None:
    if x is None:
        return None
    s = str(x).strip()
    if s == "" or s.upper() == "N/A":
        return None
    s = re.sub(r"[^0-9Xx]", "", s).upper()
    if len(s) != 8:
        return None
    return s[:4] + "-" + s[4:]

def first_nonempty(series: pd.Series):
    for v in series:
        if pd.notna(v) and str(v) != "":
            return v
    return np.nan

def quartile_rank(q: str) -> int:
    """Convert quartile (Q1–Q4) to numeric rank for comparison (Q1=1 best)."""
    if pd.isna(q):
        return 99
    q = str(q).strip().upper()
    if q in ["Q1","Q2","Q3","Q4"]:
        return int(q[1])
    return 99

# 1) Load journal_abb + journal_pub
df1 = pd.read_csv(FILE_FIRST, sep="\t", dtype=str).fillna("")
df2 = pd.read_csv(FILE_SECOND, sep="\t", dtype=str).fillna("")

# Normalize keys
df1["pubmed_issn_norm"]  = df1["pubmed_issn"].apply(lambda v: normalize_issn(v) if v != "" else None)
df1["pubmed_eissn_norm"] = df1["pubmed_eissn"].apply(lambda v: normalize_issn(v) if v != "" else None)
df2["ISSN_norm"]  = df2["ISSN"].apply(normalize_issn)
df2["eISSN_norm"] = df2["eISSN"].apply(normalize_issn)

# Deduplicate pub dataset (ISSN/eISSN keys unique)
issn_dedup = (
    df2.dropna(subset=["ISSN_norm"])
       .groupby("ISSN_norm", as_index=False)
       .agg({"Publisher": first_nonempty, JOURNAL_PUB_SOURCE: first_nonempty})
)
eissn_dedup = (
    df2.dropna(subset=["eISSN_norm"])
       .groupby("eISSN_norm", as_index=False)
       .agg({"Publisher": first_nonempty, JOURNAL_PUB_SOURCE: first_nonempty})
)
pub_by_issn   = issn_dedup.set_index("ISSN_norm")["Publisher"]
jpub_by_issn  = issn_dedup.set_index("ISSN_norm")[JOURNAL_PUB_SOURCE]
pub_by_eissn  = eissn_dedup.set_index("eISSN_norm")["Publisher"]
jpub_by_eissn = eissn_dedup.set_index("eISSN_norm")[JOURNAL_PUB_SOURCE]

# Map onto df1
df1["Publisher_issn"]    = pd.Series(df1["pubmed_issn_norm"]).map(pub_by_issn)
df1["journal_pub_issn"]  = pd.Series(df1["pubmed_issn_norm"]).map(jpub_by_issn)
df1["Publisher_eissn"]   = pd.Series(df1["pubmed_eissn_norm"]).map(pub_by_eissn)
df1["journal_pub_eissn"] = pd.Series(df1["pubmed_eissn_norm"]).map(jpub_by_eissn)

df1["Publisher"]   = df1["Publisher_issn"].combine_first(df1["Publisher_eissn"]).fillna("")
df1["journal_pub"] = df1["journal_pub_issn"].combine_first(df1["journal_pub_eissn"]).fillna("")

has_issn  = df1["Publisher_issn"].notna() | df1["journal_pub_issn"].notna()
has_eissn = df1["Publisher_eissn"].notna() | df1["journal_pub_eissn"].notna()
df1["match_pub"] = np.select([has_issn, (~has_issn) & has_eissn], ["ISSN","eISSN"], default="N/A")
# 2) Load journal_if dataset
df_if = pd.read_csv(FILE_THIRD, sep="\t", dtype=str).fillna("")

# Normalize keys
df_if["issn_norm"]  = df_if["issn"].apply(normalize_issn)
df_if["eissn_norm"] = df_if["eissn"].apply(normalize_issn)
df_if["year"] = df_if["year"].astype(int)
df_if["jif"]  = pd.to_numeric(df_if["jif"], errors="coerce")

# Deduplicate journal-year-category rows
df_if_qc = df_if.drop_duplicates(subset=["issn_norm","eissn_norm","year","category"]).copy()
df_if_qc.loc[:, "quartile_rank"] = df_if_qc["category_quartile"].apply(quartile_rank)

# Aggregate: one row per journal-year
agg = (
    df_if_qc.groupby(["issn_norm","eissn_norm","year"])
    .agg(
        jif_val=("jif","first"),   # JIF identical for same journal-year
        best_q=("quartile_rank","min")
    )
    .reset_index()
)
agg["best_quartile"] = agg["best_q"].map({1:"Q1",2:"Q2",3:"Q3",4:"Q4"}).fillna("")

# Pivot to wide format
if_wide = agg.pivot_table(
    index=["issn_norm","eissn_norm"],
    columns="year",
    values=["jif_val","best_quartile"],
    aggfunc="first"
).reset_index()

if_wide.columns = [f"{'IF' if c[0]=='jif_val' else 'q'}_{c[1]}" if c[1] else c[0] for c in if_wide.columns]
# Keep journal name + IF columns
if_by_issn = (
    if_wide.dropna(subset=["issn_norm"])
           .merge(
               df_if_qc.groupby("issn_norm")["journal"].first().reset_index(),
               on="issn_norm",
               how="left"
           )
           .groupby("issn_norm")
           .first()
)
if_by_issn.rename(columns={"journal":"journal_if"}, inplace=True)

if_by_eissn = (
    if_wide.dropna(subset=["eissn_norm"])
           .merge(
               df_if_qc.groupby("eissn_norm")["journal"].first().reset_index(),
               on="eissn_norm",
               how="left"
           )
           .groupby("eissn_norm")
           .first()
)
if_by_eissn.rename(columns={"journal":"journal_if"}, inplace=True)

# Merge IF columns by ISSN first
df1_if = df1.merge(if_by_issn, how="left", left_on="pubmed_issn_norm", right_index=True)

# For rows still unmatched, merge by eISSN
df1_if = df1_if.merge(if_by_eissn, how="left", left_on="pubmed_eissn_norm", right_index=True, suffixes=("","_eissn"))

# Fill missing IF/q/journal_if values from eISSN where ISSN was missing
for col in if_by_issn.columns:
    if f"{col}_eissn" in df1_if.columns:
        df1_if[col] = df1_if[col].combine_first(df1_if[f"{col}_eissn"])
        df1_if.drop(columns=[f"{col}_eissn"], inplace=True)

# match_if column
has_if_issn  = df1_if["pubmed_issn_norm"].isin(if_by_issn.index)
has_if_eissn = df1_if["pubmed_eissn_norm"].isin(if_by_eissn.index)
df1_if["match_if"] = np.select([has_if_issn, (~has_if_issn) & has_if_eissn], ["ISSN","eISSN"], default="N/A")

# 4) Final cleanup
drop_cols = ["pubmed_issn_norm","pubmed_eissn_norm","Publisher_issn","Publisher_eissn","journal_pub_issn","journal_pub_eissn","issn_norm","eissn_norm","best_q"]
df_out = df1_if.drop(columns=[c for c in drop_cols if c in df1_if.columns])

df_out.to_csv(FILE_OUT, sep="\t", index=False)
print(f"Done. Wrote: {FILE_OUT} | Rows: {len(df_out):,}")