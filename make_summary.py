"""data/ の生データ（15分ごとのCSV）をまとめて，Excelで見やすい集計CSVを summary/ に作る．

作るファイル
  summary/all_data.csv          全期間を1ファイルにまとめたもの（日本語の列名，状態つき）
  summary/bikes_wide.csv        行＝日時，列＝ステーションの貸出可能台数（Excelでグラフにしやすい形）
  summary/daily_by_station.csv  ステーション×日ごとの集計（満車・空車の回数と割合など）
"""
import glob
import os

import pandas as pd

DATA_DIR = "data"
OUT_DIR = "summary"

# 区の対応（scrape.py の TARGET_IDS と同じ25箇所）
WARD = {
    **{i: "中区" for i in ["12160", "24522", "15109", "24429", "16161", "16597",
                          "13122", "17524", "17802", "24332", "4045", "15110"]},
    **{i: "西区" for i in ["16613", "22685", "17577", "9788", "16720", "6569", "15389",
                          "24686", "19967", "27703", "10879", "15124", "17801"]},
}


def load_raw():
    files = sorted(glob.glob(f"{DATA_DIR}/yokohama_stations_15min_log_raw_*.csv"))
    if not files:
        raise SystemExit("data/ に生データがありません")
    df = pd.concat(
        [pd.read_csv(f, dtype={"station_id": str}, encoding="utf-8-sig") for f in files],
        ignore_index=True,
    )
    df["fetch_time"] = pd.to_datetime(df["fetch_time"], format="%Y/%m/%d %H:%M")
    df["missing"] = df["missing"].astype(str).str.lower().eq("true")
    # 同じ回・同じステーションの重複（再実行などで起こりうる）を除く
    df = df.drop_duplicates(subset=["fetch_time", "station_id"], keep="last")
    return df.sort_values(["fetch_time", "station_id"]).reset_index(drop=True)


def add_status(df):
    def status(r):
        if r["missing"]:
            return "欠損"
        if r["is_renting"] is False or str(r["is_renting"]).lower() == "false":
            return "休止"
        if r["num_docks_available"] == 0:
            return "満車"
        if r["num_bikes_available"] == 0:
            return "空車"
        return "通常"

    df["status"] = df.apply(status, axis=1)
    df["ward"] = df["station_id"].map(WARD).fillna("")
    # 名前は欠損行で空になるので，同じIDの他の行から補う
    names = df.dropna(subset=["name"]).groupby("station_id")["name"].last()
    caps = df.dropna(subset=["capacity"]).groupby("station_id")["capacity"].last()
    df["name"] = df["name"].fillna(df["station_id"].map(names))
    df["capacity"] = df["capacity"].fillna(df["station_id"].map(caps))
    return df


def write_all_data(df):
    out = pd.DataFrame({
        "日時": df["fetch_time"].dt.strftime("%Y/%m/%d %H:%M"),
        "日付": df["fetch_time"].dt.strftime("%Y/%m/%d"),
        "時刻": df["fetch_time"].dt.strftime("%H:%M"),
        "曜日": df["fetch_time"].dt.dayofweek.map(dict(enumerate("月火水木金土日"))),
        "区": df["ward"],
        "ステーションID": df["station_id"],
        "ステーション名": df["name"],
        "容量": df["capacity"],
        "貸出可能台数": df["num_bikes_available"],
        "空き枠": df["num_docks_available"],
        "貸出不可台数": df["unavailable"],
        "状態": df["status"],
        "ステーション最終更新": df["last_reported_jst"],
    })
    out.to_csv(f"{OUT_DIR}/all_data.csv", index=False, encoding="utf-8-sig")


def write_bikes_wide(df):
    ok = df[~df["missing"]].copy()
    ok["列名"] = ok["ward"] + "_" + ok["name"].astype(str) + "(" + ok["station_id"] + ")"
    wide = ok.pivot_table(index="fetch_time", columns="列名",
                          values="num_bikes_available", aggfunc="last")
    # 列は 区→ステーション名 の順
    wide = wide[sorted(wide.columns)]
    wide.index = wide.index.strftime("%Y/%m/%d %H:%M")
    wide.index.name = "日時"
    wide.to_csv(f"{OUT_DIR}/bikes_wide.csv", encoding="utf-8-sig")


def write_daily(df):
    df = df.copy()
    df["date"] = df["fetch_time"].dt.strftime("%Y/%m/%d")
    g = df.groupby(["date", "ward", "station_id", "name"], dropna=False)
    daily = g.agg(
        容量=("capacity", "last"),
        取得回数=("status", "size"),
        欠損回数=("status", lambda s: (s == "欠損").sum()),
        満車回数=("status", lambda s: (s == "満車").sum()),
        空車回数=("status", lambda s: (s == "空車").sum()),
        平均貸出可能台数=("num_bikes_available", "mean"),
        最小貸出可能台数=("num_bikes_available", "min"),
        最大貸出可能台数=("num_bikes_available", "max"),
        平均貸出不可台数=("unavailable", "mean"),
        最大貸出不可台数=("unavailable", "max"),
    ).reset_index()
    valid = (daily["取得回数"] - daily["欠損回数"]).replace(0, pd.NA)
    daily["満車率(%)"] = (daily["満車回数"] / valid * 100).astype(float).round(1)
    daily["空車率(%)"] = (daily["空車回数"] / valid * 100).astype(float).round(1)
    for c in ["平均貸出可能台数", "平均貸出不可台数"]:
        daily[c] = daily[c].astype(float).round(2)
    daily = daily.rename(columns={"date": "日付", "ward": "区",
                                  "station_id": "ステーションID", "name": "ステーション名"})
    daily.sort_values(["日付", "区", "ステーションID"]).to_csv(
        f"{OUT_DIR}/daily_by_station.csv", index=False, encoding="utf-8-sig")


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    df = add_status(load_raw())
    write_all_data(df)
    write_bikes_wide(df)
    write_daily(df)
    print(f"✅ 集計CSVを作成しました（{df['fetch_time'].nunique()} 回分，{len(df)} 行）")


if __name__ == "__main__":
    main()
