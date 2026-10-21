import os
import sys
import json
import time
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests

SAVE_DIR = "data"
os.makedirs(SAVE_DIR, exist_ok=True)

# 監視対象：横浜市中区・西区のHELLO CYCLING単独ポート
# （ドコモとの共同ポート「001.○○」形式を除外，容量5〜40，2026/10/21 23:20時点で貸出・返却とも稼働中）
TARGET_IDS = [
    # 中区（12箇所）
    '12160', '24522', '15109', '24429', '16161', '16597',
    '13122', '17524', '17802', '24332', '4045', '15110',
    # 西区（13箇所）
    '16613', '22685', '17577', '9788', '16720', '6569', '15389',
    '24686', '19967', '27703', '10879', '15124', '17801',
]

STATUS_URL = "https://api-public.odpt.org/api/v4/gbfs/hellocycling/station_status.json"
INFO_URL = "https://api-public.odpt.org/api/v4/gbfs/hellocycling/station_information.json"

JST = timezone(timedelta(hours=9), "JST")

COLUMNS_TO_SAVE = [
    "fetch_time", "fetch_time_iso", "api_last_updated",
    "station_id", "name", "capacity",
    "is_installed", "is_renting", "is_returning",
    "num_bikes_available", "num_docks_available",
    "unavailable",  # 容量 −（貸出可能台数＋空き枠）．故障・電池切れ・予約中などで貸せない車両の目安
    "last_reported", "last_reported_jst",
    "vehicle_docks_available", "vehicle_types_available",
    "missing",
]


def get_json(url, retries=3, wait=10):
    """タイムアウトとリトライつきでJSONを取得する．"""
    last_err = None
    for i in range(retries):
        try:
            r = requests.get(url, timeout=30)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            last_err = e
            print(f"  取得失敗（{i + 1}/{retries}回目）: {e}")
            if i < retries - 1:
                time.sleep(wait)
    raise last_err


def to_jst(ts):
    """UNIX時刻を日本時間の文字列に変換する．変換できなければNoneを返す．"""
    try:
        return datetime.fromtimestamp(int(ts), JST).strftime("%Y/%m/%d %H:%M:%S")
    except (TypeError, ValueError):
        return None


def normalize_id(series):
    """station_idから数字部分だけを取り出す．"""
    return series.astype(str).str.extract(r"(\d+)")[0]


def main():
    now = datetime.now(JST)
    csv_file = f"{SAVE_DIR}/yokohama_stations_15min_log_raw_{now:%Y%m%d}.csv"

    print("▶ データ取得開始...")

    # 1. 稼働状況（これが取れなければ失敗扱い）
    status_json = get_json(STATUS_URL)
    stations = status_json.get("data", {}).get("stations")
    if not stations:
        raise RuntimeError("station_status にステーション一覧が含まれていません")

    df = pd.DataFrame(stations)
    df["station_id"] = normalize_id(df["station_id"])
    df = df[df["station_id"].isin(TARGET_IDS)].copy()

    # 2. 名前と容量（GBFS 2.3では容量は vehicle_capacity に文字列で入っている）
    try:
        info = pd.DataFrame(get_json(INFO_URL)["data"]["stations"])
        info["station_id"] = normalize_id(info["station_id"])
        cap_col = "capacity" if "capacity" in info.columns else "vehicle_capacity"
        info["capacity"] = pd.to_numeric(info.get(cap_col), errors="coerce")
        df = df.merge(info[["station_id", "name", "capacity"]], on="station_id", how="left")
    except Exception as e:
        print(f"⚠️ station_information の取得に失敗（名前・容量は空欄で保存）: {e}")

    # 3. 貸せない車両の目安（容量 −（貸出可能台数＋空き枠））
    if "capacity" in df.columns:
        df["unavailable"] = df["capacity"] - (
            pd.to_numeric(df["num_bikes_available"], errors="coerce")
            + pd.to_numeric(df["num_docks_available"], errors="coerce")
        )

    # 4. APIに存在しなかったIDを欠損行として残す
    df["missing"] = False
    missing_ids = sorted(set(TARGET_IDS) - set(df["station_id"]))
    if missing_ids:
        print(f"⚠️ APIに存在しないID（{len(missing_ids)}件）: {missing_ids}")
        df_missing = pd.DataFrame({"station_id": missing_ids, "missing": True})
        df = pd.concat([df, df_missing], ignore_index=True)

    # 5. 時刻情報を付与する
    df["fetch_time"] = now.strftime("%Y/%m/%d %H:%M")
    df["fetch_time_iso"] = now.isoformat(timespec="seconds")
    df["api_last_updated"] = to_jst(status_json.get("last_updated"))
    if "last_reported" in df.columns:
        df["last_reported_jst"] = df["last_reported"].apply(to_jst)

    # 6. リストや辞書の列はJSON文字列にして保存する
    for col in ["vehicle_docks_available", "vehicle_types_available"]:
        if col in df.columns:
            df[col] = df[col].apply(
                lambda v: json.dumps(v, ensure_ascii=False) if isinstance(v, (list, dict)) else v
            )

    # 7. 列を揃えて追記する
    for col in COLUMNS_TO_SAVE:
        if col not in df.columns:
            df[col] = None
    df_final = df[COLUMNS_TO_SAVE].sort_values("station_id")

    write_header = not os.path.exists(csv_file)
    df_final.to_csv(csv_file, mode="a", header=write_header, index=False, encoding="utf-8-sig")

    n_ok = int((~df_final["missing"].astype(bool)).sum())
    print(f"✅ 保存成功（取得 {n_ok} 箇所 / 欠損 {len(missing_ids)} 箇所）→ {csv_file}")

    # 対象が1件も取れなかった場合は，異常としてジョブを失敗させる
    if n_ok == 0:
        print("❌ 対象ステーションが1件も取得できませんでした")
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"❌ エラー発生: {e}")
        sys.exit(1)
