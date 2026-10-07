"""Build the integrated province-year flood dataset (Thailand, 2019-2024).

Sources (raw files in data/raw/):
  ddpm/25xx.xlsx          DDPM yearly flood statistics (schema differs every year)
  rain/spatial-rain-hii.csv  HII / ThaiWater monthly areal rainfall by province
  rain/avg30y_province.csv   HII 30-year (1991-2020) monthly normal by province
  dopa/cXX/*.xls          DOPA registered population (HTML tables saved as .xls)
  ldd/sumflood.csv        LDD recurrent-flood area by sub-district (validation)
  geo/provinces.geojson   77 province polygons

Outputs (data/processed/):
  ddpm_long.csv           harmonised province-year flood impact
  province_year.csv       DDPM + rain + population joined by province & year
  province_features.csv   one row per province (77) for clustering
  ldd_recurrent.csv       recurrent-flood share per province
  data_log.json           counts used in the EDA / preprocessing slides
"""
import io
import json
import re
from pathlib import Path

import geopandas as gpd
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
OUT = ROOT / "data" / "processed"
OUT.mkdir(parents=True, exist_ok=True)

YEARS_BE = [2562, 2563, 2564, 2565, 2566, 2567]
BE2CE = {y: y - 543 for y in YEARS_BE}
LOG = {}

# --------------------------------------------------------------------------
# Province name normalisation
# --------------------------------------------------------------------------
GEO = gpd.read_file(RAW / "geo" / "provinces.geojson")
GEO["pro_code"] = GEO["pro_code"].astype(int)
NAME2CODE = dict(zip(GEO["pro_th"], GEO["pro_code"]))
ALIASES = {
    "กรุงเทพฯ": "กรุงเทพมหานคร", "กทม": "กรุงเทพมหานคร", "กทม.": "กรุงเทพมหานคร",
    "กรุงเทพ": "กรุงเทพมหานคร",
}


def norm_province(name):
    """Map any spelling of a province to its numeric code (None if not a province)."""
    if not isinstance(name, str):
        return None
    s = re.sub(r"[​﻿\s]+", "", name)          # zero-width / all whitespace
    s = re.sub(r"^(จังหวัด|จ\.)", "", s)
    if "รวม" in s or "ผลรวม" in s:                       # subtotal / grand-total rows
        return None
    s = ALIASES.get(s, s)
    return NAME2CODE.get(s)


def to_num(s):
    """'1,470' / ' ' / '-' / NaN -> float (blank and '-' become NaN)."""
    return pd.to_numeric(
        s.astype(str).str.replace(",", "", regex=False).str.strip()
         .replace({"": np.nan, "-": np.nan, "nan": np.nan, "None": np.nan}),
        errors="coerce")


# --------------------------------------------------------------------------
# 1) DDPM: harmonise 6 different yearly schemas to 12 common attributes
# --------------------------------------------------------------------------
# Each harmonised attribute = sum of the matching raw columns (matched by Thai
# keyword) so the detailed 2019-2022 schema maps onto the coarse 2023-2024 one.
GROUPS = {
    "affected_people":   ["ราษฎร (คน)", "ราษฎร(คน)", "ราษฎร \n(คน)", "ผู้ได้รับผลกระทบ (คน)", "ผู้ได้รับผลกระทบ\n(คน)", "Affected People"],
    "affected_hh":       ["ราษฎร (ครัวเรือน)", "ราษฎร(ครัวเรือน)", "ราษฎร \n(ครัวเรือน)", "ผู้ได้รับผลกระทบ (ครัวเรือน)", "ผู้ได้รับผลกระทบ\n(ครัวเรือน)", "Affected Households"],
    "evacuated":         ["อพยพ (คน)", "อพยพ(คน)", "อพยพ \n(คน)", "ผู้อพยพ (คน)", "ผู้อพยพ\n(คน)", "Evacuees"],
    "deaths":            ["เสียชีวิต", "fatalities_count"],
    "injured_missing":   ["บาดเจ็บ", "สูญหาย", "injuries_count", "Missing Persons"],
    "housing":           ["บ้านพักอาศัย", "บ้านเรือนบางส่วน", "อาคารชุด", "อาคารพาณิชย์", "ด้านที่อยู่อาศัย", "Housing"],
    "agri_rai":          ["นาข้าว", "พืชสวน", "พืชไร่", "พืชต้นทุนสูง(ไร่)", "พืชต้นทุนสูง \n(ไร่)", "ด้านการเกษตร", "Agriculture"],
    "livestock":         ["โค", "กระบือ", "แพะ", "แกะ", "ไก่", "เป็ด", "ด้านปศุสัตว์", "ด้าน\nปศุสัตว์", "Livestock"],
    "fishery":           ["บ่อปลา", "บ่อกุ้ง", "กระชังปลา", "ด้านประมง", "Fisheries"],
    "transport":         ["ถนน", "สะพาน", "ด้านคมนาคม", "ด้าน\nคมนาคม", "Transportation"],
    "public_sites":      ["วัด/มัสยิด", "โรงเรียน", "โรงพยาบาล", "สถานที่ราชการ", "ไฟฟ้า", "ประปา",
                          "ทำนบ", "ฝาย", "เหมือง", "พนังกั้นน้ำ", "บ่อน้ำ/อ่างเก็บน้ำ", "ท่อระบายน้ำ",
                          "ด้านสาธารณสุข", "ด้าน\nสาธารณสุข", "ศาสนา", "สถานศึกษา", "สาธารณูปโภค", "สถานที่\nราชการ", "ด้านสถานที่ราชการ",
                          "Health", "Religion", "Education", "Utilities", "Government Sites"],
}
# Columns that must never be captured by the keyword match above
EXCLUDE = ["เขต", "แขวง", "พืชต้นทุนสูง(ต้น)", "ยานยนต์", "ไก่ชน"]
HARM = list(GROUPS) + ["n_districts", "n_villages"]


def pick(cols, keys):
    out = []
    for c in cols:
        cs = str(c)
        if "มูลค่า" in cs or "THB" in cs:      # monetary columns exist only in 2019
            continue
        flat = cs.replace("\n", "").replace(" ", "")
        for k in keys:
            kf = k.replace("\n", "").replace(" ", "")
            if flat.startswith(kf) or (len(kf) > 4 and kf in flat):
                if not any(e.replace(" ", "") == flat for e in EXCLUDE):
                    out.append(c)
                break
    return out


def harmonise(df, prov_col, district_col=None, village_col=None, n_villages_from_rows=False):
    df = df.copy()
    df["pro_code"] = df[prov_col].map(norm_province)
    dropped = df["pro_code"].isna().sum()
    df = df[df["pro_code"].notna()]
    res = pd.DataFrame({"pro_code": df["pro_code"].astype(int)})
    used = {}
    for g, keys in GROUPS.items():
        cols = [c for c in pick(df.columns, keys) if c not in sum(used.values(), [])]
        used[g] = cols
        res[g] = sum((to_num(df[c]).fillna(0) for c in cols), pd.Series(0.0, index=df.index)) if cols else np.nan
    if district_col is not None:
        res["district_key"] = df[district_col].astype(str).str.strip()
    if village_col is not None:
        res["n_villages"] = to_num(df[village_col]).fillna(0)
    if n_villages_from_rows:
        res["n_villages"] = 1
    return res, used, int(dropped)


def load_ddpm():
    frames, meta = [], {}

    # 2562 (2019): already one row per province, English+Thai header rows
    d = pd.read_excel(RAW / "ddpm/2562.xlsx", sheet_name="flood2562cut", header=1)
    r, used, dr = harmonise(d, "จังหวัด", village_col="จำนวนหมู่บ้าน")
    r["n_districts"] = to_num(d.loc[r.index, "อำเภอ"])
    frames.append(r.assign(year_be=2562)); meta[2562] = dict(rows=len(d), dropped_total_rows=dr, level="province")

    # 2563 (2020): one row per district-event, subtotal rows "X รวม"
    d = pd.read_excel(RAW / "ddpm/2563.xlsx", sheet_name="ความเสียหายรวม", header=0)
    r, used, dr = harmonise(d, "จังหวัด", district_col="อำเภอ", village_col="จำนวนหมู่บ้านและชุมชน")
    frames.append(r.assign(year_be=2563)); meta[2563] = dict(rows=len(d), dropped_total_rows=dr, level="district-event")

    # 2564 (2021): two-row merged header, district-event rows + subtotals
    d = pd.read_excel(RAW / "ddpm/2564.xlsx", sheet_name="ความเสียหายรวม", header=1)
    r, used, dr = harmonise(d, "จังหวัด", district_col="อำเภอ", village_col="หมู่บ้าน/\nชุมชน")
    frames.append(r.assign(year_be=2564)); meta[2564] = dict(rows=len(d), dropped_total_rows=dr, level="district-event")

    # 2565 (2022): two-row header, high-value crops given in trees AND rai
    d = pd.read_excel(RAW / "ddpm/2565.xlsx", sheet_name="อำเภอ+ทั้งหมด", header=1)
    r, used, dr = harmonise(d, "จังหวัด", district_col="อำเภอ", village_col="หมู่บ้าน/ชุมชน")
    frames.append(r.assign(year_be=2565)); meta[2565] = dict(rows=len(d), dropped_total_rows=dr, level="district-event")

    # 2566 (2023): new coarse damage scheme, province summary with 3-row header
    raw = pd.read_excel(RAW / "ddpm/2566.xlsx", sheet_name="พื้นที่เกิดรายจังหวัด", header=None)
    top = raw.iloc[2].ffill().fillna("").astype(str)
    sub = raw.iloc[3].fillna("").astype(str)
    cols = [s if s else t for t, s in zip(top, sub)]
    d = raw.iloc[4:].copy(); d.columns = cols
    r, used, dr = harmonise(d, "จังหวัด", village_col="จำนวนหมู่ที่\n/ชุมชน")
    r["n_districts"] = to_num(d.loc[r.index, "จำนวนอำเภอ/เขต"])
    frames.append(r.assign(year_be=2566)); meta[2566] = dict(rows=len(d), dropped_total_rows=dr, level="province")

    # 2567 (2024): one row per affected village (~35k rows), English header
    d = pd.read_excel(RAW / "ddpm/2567.xlsx", sheet_name=1, header=0).iloc[1:]   # row 1 = Thai labels
    d.columns = [str(c).strip() for c in d.columns]
    r, used, dr = harmonise(d, "Province", district_col="District Code", n_villages_from_rows=True)
    frames.append(r.assign(year_be=2567)); meta[2567] = dict(rows=len(d), dropped_total_rows=dr, level="village")

    long = pd.concat(frames, ignore_index=True)
    # district-level years: number of distinct districts per province
    nd = (long.dropna(subset=["district_key"]).groupby(["year_be", "pro_code"])["district_key"].nunique()
          .rename("nd_from_rows"))
    agg = long.drop(columns=["district_key"]).groupby(["year_be", "pro_code"]).sum(min_count=1)
    agg = agg.join(nd)
    agg["n_districts"] = agg["n_districts"].where(agg["n_districts"] > 0, agg["nd_from_rows"])
    agg = agg.drop(columns="nd_from_rows").reset_index()
    agg["year"] = agg["year_be"].map(BE2CE)
    LOG["ddpm_meta"] = {str(k): v for k, v in meta.items()}
    return agg


# --------------------------------------------------------------------------
# 2) Rainfall (HII / ThaiWater)
# --------------------------------------------------------------------------
def load_rain():
    r = pd.read_csv(RAW / "rain/spatial-rain-hii.csv", encoding="utf-8-sig")
    bad = r[(r.MinRain > r.AvgRain) | (r.AvgRain > r.MaxRain)]
    LOG["rain_inconsistent_rows"] = {int(k): int(v) for k, v in bad.groupby("YEAR").size().items()}
    r = r[r.YEAR.between(2019, 2024)]
    yr = r.groupby(["YEAR", "PROV_ID"]).agg(rain_annual_mm=("AvgRain", "sum"),
                                             rain_max_month_mm=("AvgRain", "max")).reset_index()
    n = pd.read_csv(RAW / "rain/avg30y_province.csv", encoding="utf-8-sig")
    normal = n.groupby("PROV_CODE")["MEAN"].sum().rename("rain_normal_mm")
    yr = yr.merge(normal, left_on="PROV_ID", right_index=True, how="left")
    yr["rain_anomaly_pct"] = 100 * (yr.rain_annual_mm / yr.rain_normal_mm - 1)
    return yr.rename(columns={"YEAR": "year", "PROV_ID": "pro_code"})


# --------------------------------------------------------------------------
# 3) Population (DOPA) – the .xls files are actually UTF-8 HTML tables
# --------------------------------------------------------------------------
def load_pop():
    rows = []
    for f in sorted((RAW / "dopa").glob("c*/*.xls")):
        t = pd.read_html(io.StringIO(f.read_text(encoding="utf-8")), header=0)[0]
        t.columns = [c.replace("﻿", "") for c in t.columns]
        t["year_be"] = 2500 + int(re.search(r"c(\d\d)", str(f.parent)).group(1))
        rows.append(t)
    p = pd.concat(rows)
    p = p[p["รหัสจังหวัด"] != 0]
    p = p.rename(columns={"รหัสจังหวัด": "pro_code", "จำนวนประชากรทั้งหมด": "population",
                          "จำนวนบ้าน (หลังคาเรือน)": "houses"})
    p["year"] = p["year_be"] - 543
    return p[["year", "pro_code", "population", "houses"]]


# --------------------------------------------------------------------------
# 4) LDD recurrent flood (validation reference)
# --------------------------------------------------------------------------
def load_ldd():
    l = pd.read_csv(RAW / "ldd/sumflood.csv", encoding="utf-8-sig")
    l.columns = ["prov", "amphoe", "tambon", "f8_10", "f4_7", "f_lt3", "area_rai", "update"]
    l["prov"] = l["prov"].ffill()                                 # merged cells
    l["pro_code"] = l["prov"].map(norm_province)
    for c in ["f8_10", "f4_7", "f_lt3", "area_rai"]:
        l[c] = to_num(l[c]).fillna(0)
    g = l.groupby("pro_code")[["f8_10", "f4_7", "f_lt3", "area_rai"]].sum()
    area = GEO.set_index("pro_code")["area_sqkm"] * 625                # 1 km2 = 625 rai
    g = g.reindex(GEO["pro_code"]).fillna(0)
    g["recurrent_rai"] = g["f8_10"] + g["f4_7"]                          # >= 4 floods in 10 years
    g["recurrent_share"] = g["recurrent_rai"] / area.reindex(g.index)
    g["flood_any_share"] = g["area_rai"] / area.reindex(g.index)
    LOG["ldd_provinces_with_data"] = int((g["area_rai"] > 0).sum())
    return g.reset_index()


# --------------------------------------------------------------------------
def main():
    ddpm = load_ddpm()
    rain = load_rain()
    pop = load_pop()
    ldd = load_ldd()

    # full 77 x 6 grid: a province that is absent from a DDPM year had no declared flood
    grid = pd.MultiIndex.from_product([range(2019, 2025), GEO["pro_code"]], names=["year", "pro_code"]).to_frame(index=False)
    py = grid.merge(ddpm.drop(columns="year_be"), on=["year", "pro_code"], how="left")
    py["flood_reported"] = py["affected_hh"].notna() | py["n_villages"].notna()
    LOG["province_years_reported"] = int(py["flood_reported"].sum())

    # zero-means-missing: households reported but people = 0 (or vice versa)
    hh_ratio = (py.affected_people / py.affected_hh).replace([np.inf, 0], np.nan)
    ratio = hh_ratio.groupby(py.pro_code).transform("median").fillna(hh_ratio.median())
    m1 = (py.affected_people.fillna(0) == 0) & (py.affected_hh > 0)
    m2 = (py.affected_hh.fillna(0) == 0) & (py.affected_people > 0)
    py.loc[m1, "affected_people"] = (py.loc[m1, "affected_hh"] * ratio[m1]).round()
    py.loc[m2, "affected_hh"] = (py.loc[m2, "affected_people"] / ratio[m2]).round()
    LOG["imputed_people_from_households"] = int(m1.sum())
    LOG["imputed_households_from_people"] = int(m2.sum())
    LOG["median_people_per_household"] = float(hh_ratio.median())

    # non-reported years are true zeros (no flood declared)
    num = [c for c in HARM if c in py]
    py[num] = py[num].fillna(0)

    py = py.merge(rain, on=["year", "pro_code"], how="left").merge(pop, on=["year", "pro_code"], how="left")
    py = py.merge(GEO[["pro_code", "pro_th", "pro_en", "reg_royin"]], on="pro_code")
    py.to_csv(OUT / "province_year.csv", index=False)
    LOG["province_year_shape"] = list(py.shape)

    # ---- one row per province (77) ------------------------------------------
    g = py.groupby("pro_code")
    f = pd.DataFrame({
        "years_flooded": g["flood_reported"].sum(),
        "pop_mean": g["population"].mean(),
        "rain_annual_mm": g["rain_annual_mm"].mean(),
        "rain_max_month_mm": g["rain_max_month_mm"].mean(),
        "rain_anomaly_pct": g["rain_anomaly_pct"].mean(),
    })
    tot = g[num].sum()
    f = f.join(tot.add_prefix("tot_"))
    # intensity per 1,000 residents (removes "big province = big numbers" effect)
    for c in ["affected_people", "evacuated", "deaths", "injured_missing", "housing", "public_sites"]:
        f[f"{c}_per_1k"] = 1000 * tot[c] / 6 / f["pop_mean"]
    f["agri_rai_per_km2"] = tot["agri_rai"] / 6 / GEO.set_index("pro_code")["area_sqkm"]
    f["villages_per_year"] = tot["n_villages"] / 6
    f = f.join(GEO.set_index("pro_code")[["pro_th", "pro_en", "reg_royin", "area_sqkm"]])
    f = f.join(ldd.set_index("pro_code")[["recurrent_share", "flood_any_share", "recurrent_rai"]])
    f.reset_index().to_csv(OUT / "province_features.csv", index=False)
    ldd.to_csv(OUT / "ldd_recurrent.csv", index=False)
    ddpm.to_csv(OUT / "ddpm_long.csv", index=False)

    LOG["n_raw_ddpm_attributes"] = {"2562": 49, "2563": 60, "2564": 59, "2565": 56, "2566": 66, "2567": 44}
    LOG["n_harmonised_attributes"] = len(num)
    (OUT / "data_log.json").write_text(json.dumps(LOG, ensure_ascii=False, indent=2))
    print(json.dumps(LOG, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
