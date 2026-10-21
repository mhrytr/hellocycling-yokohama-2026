"""make_summary.py が作った summary/*.csv から，罫線・色つきの見やすいExcel（summary/summary.xlsx）を作る．

シート
  最新の状況   各ステーションの最新の台数と状態
  日別集計     ステーション×日ごとの満車・空車の回数と割合
  台数推移     直近7日の貸出可能台数（行＝日時，列＝ステーション）．色の濃さで台数が分かる
  直近7日の全データ  15分ごとの全記録（フィルターつき）
  説明         列と色の意味
"""
from datetime import timedelta

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT_DIR = "summary"
RECENT_DAYS = 7

FONT_NAME = "Yu Gothic"
F_BODY = Font(name=FONT_NAME, size=10)
F_HEAD = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT_NAME, size=14, bold=True)
F_NOTE = Font(name=FONT_NAME, size=9, color="595959")

FILL_HEAD = PatternFill("solid", fgColor="1F4E78")
FILL_NAKA = PatternFill("solid", fgColor="EAF1FB")   # 中区の行の背景
FILL_NISHI = PatternFill("solid", fgColor="FFFFFF")
FILL_FULL = PatternFill("solid", fgColor="F8CBAD")   # 満車
FILL_EMPTY = PatternFill("solid", fgColor="BDD7EE")  # 空車
FILL_MISS = PatternFill("solid", fgColor="D9D9D9")   # 欠損
FILL_STOP = PatternFill("solid", fgColor="FFE699")   # 休止

THIN = Side(style="thin", color="A6A6A6")
MED = Side(style="medium", color="1F4E78")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")

STATUS_FILL = {"満車": FILL_FULL, "空車": FILL_EMPTY, "欠損": FILL_MISS, "休止": FILL_STOP}


def write_table(ws, df, start_row, widths, number_formats=None, text_cols=()):
    """見出し行と本体を書き，罫線と書式をつける．終わりの行番号を返す．"""
    number_formats = number_formats or {}
    for j, col in enumerate(df.columns, 1):
        c = ws.cell(row=start_row, column=j, value=col)
        c.font, c.fill, c.alignment, c.border = F_HEAD, FILL_HEAD, CENTER, BORDER
    ws.row_dimensions[start_row].height = 30

    for i, row in enumerate(df.itertuples(index=False), start_row + 1):
        for j, (col, v) in enumerate(zip(df.columns, row), 1):
            if pd.isna(v):
                v = None
            c = ws.cell(row=i, column=j, value=v)
            c.font, c.border = F_BODY, BORDER
            c.alignment = LEFT if col in text_cols else RIGHT
            if col in number_formats:
                c.number_format = number_formats[col]

    for j, col in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(j)].width = widths.get(col, 10)
    return start_row + len(df)


def outline(ws, r1, c1, r2, c2):
    """表の外枠を太線にする．"""
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            b = cell.border
            cell.border = Border(
                left=MED if c == c1 else b.left, right=MED if c == c2 else b.right,
                top=MED if r == r1 else b.top, bottom=MED if r == r2 else b.bottom,
            )


def title(ws, text, note):
    ws["A1"] = text
    ws["A1"].font = F_TITLE
    ws["A2"] = note
    ws["A2"].font = F_NOTE
    ws.sheet_view.showGridLines = False


def sheet_latest(wb, df):
    ws = wb.active
    ws.title = "最新の状況"
    latest_time = df["日時"].max()
    last = df[df["日時"] == latest_time].sort_values(["区", "ステーションID"])
    cols = ["区", "ステーションID", "ステーション名", "容量", "貸出可能台数", "空き枠", "貸出不可台数", "状態"]
    t = last[cols]
    title(ws, f"最新の状況（{latest_time} 時点）",
          "色：満車＝赤，空車＝青，欠損＝灰，休止＝黄．貸出可能台数のバーは容量に対する割合の目安．")
    end = write_table(ws, t, 4,
                      {"区": 6, "ステーションID": 13, "ステーション名": 44, "容量": 7,
                       "貸出可能台数": 14, "空き枠": 8, "貸出不可台数": 13, "状態": 8},
                      text_cols={"区", "ステーションID", "ステーション名", "状態"})
    for i, (_, r) in enumerate(t.iterrows(), 5):
        base = FILL_NAKA if r["区"] == "中区" else FILL_NISHI
        for j in range(1, len(cols) + 1):
            ws.cell(row=i, column=j).fill = base
        st = ws.cell(row=i, column=8)
        st.alignment = CENTER
        if r["状態"] in STATUS_FILL:
            st.fill = STATUS_FILL[r["状態"]]
            st.font = Font(name=FONT_NAME, size=10, bold=True)
        if r["貸出不可台数"] and r["貸出不可台数"] > 0:
            ws.cell(row=i, column=7).font = Font(name=FONT_NAME, size=10, bold=True, color="C00000")
    ws.conditional_formatting.add(
        f"E5:E{end}", DataBarRule(start_type="num", start_value=0, end_type="max", color="70AD47"))
    outline(ws, 4, 1, end, len(cols))

    # 集計（数式）
    s = end + 2
    ws.cell(row=s, column=3, value="状態の内訳").font = Font(name=FONT_NAME, size=10, bold=True)
    for k, name in enumerate(["通常", "満車", "空車", "休止", "欠損"], 1):
        a = ws.cell(row=s + k, column=3, value=name)
        b = ws.cell(row=s + k, column=4, value=f'=COUNTIF($H$5:$H${end},C{s + k})')
        for c in (a, b):
            c.font, c.border = F_BODY, BORDER
        if name in STATUS_FILL:
            a.fill = STATUS_FILL[name]
    ws.freeze_panes = "A5"


def sheet_daily(wb, daily):
    ws = wb.create_sheet("日別集計")
    title(ws, "日別集計（ステーション×日）",
          "満車率・空車率＝満車（空車）回数 ÷（取得回数 − 欠損回数）．赤・青が濃いほど割合が高い．")
    cols = ["日付", "区", "ステーションID", "ステーション名", "容量", "取得回数", "欠損回数",
            "満車回数", "空車回数", "平均貸出可能台数", "最小貸出可能台数", "最大貸出可能台数",
            "平均貸出不可台数", "最大貸出不可台数"]
    t = daily[cols].sort_values(["日付", "区", "ステーションID"], ascending=[False, True, True])
    end = write_table(ws, t, 4,
                      {"日付": 11, "区": 6, "ステーションID": 13, "ステーション名": 44},
                      number_formats={"平均貸出可能台数": "0.00", "平均貸出不可台数": "0.00"},
                      text_cols={"日付", "区", "ステーションID", "ステーション名"})
    # 割合は数式で計算する
    for j, name in ((15, "満車率"), (16, "空車率")):
        c = ws.cell(row=4, column=j, value=name)
        c.font, c.fill, c.alignment, c.border = F_HEAD, FILL_HEAD, CENTER, BORDER
        ws.column_dimensions[get_column_letter(j)].width = 9
        src = "H" if name == "満車率" else "I"
        for i in range(5, end + 1):
            c = ws.cell(row=i, column=j, value=f"=IF(F{i}-G{i}=0,\"\",{src}{i}/(F{i}-G{i}))")
            c.number_format, c.font, c.border, c.alignment = "0.0%", F_BODY, BORDER, RIGHT
    # 日付が変わるところに太い線を引く
    prev = None
    for i in range(5, end + 1):
        d = ws.cell(row=i, column=1).value
        if prev is not None and d != prev:
            for j in range(1, 17):
                c = ws.cell(row=i, column=j)
                b = c.border
                c.border = Border(left=b.left, right=b.right, top=MED, bottom=b.bottom)
        prev = d
    ws.conditional_formatting.add(f"O5:O{end}", ColorScaleRule(
        start_type="num", start_value=0, start_color="FFFFFF", end_type="num", end_value=1, end_color="F4B183"))
    ws.conditional_formatting.add(f"P5:P{end}", ColorScaleRule(
        start_type="num", start_value=0, start_color="FFFFFF", end_type="num", end_value=1, end_color="9DC3E6"))
    ws.conditional_formatting.add(f"G5:G{end}", FormulaRule(formula=["G5>0"], fill=FILL_MISS))
    outline(ws, 4, 1, end, 16)
    ws.auto_filter.ref = f"A4:P{end}"
    ws.freeze_panes = "E5"


def sheet_wide(wb, df, info):
    ws = wb.create_sheet("台数推移")
    title(ws, f"貸出可能台数の推移（直近{RECENT_DAYS}日）",
          "数字は貸出可能台数．0（空車）は青，台数が多いほど緑が濃い．空欄は欠損．")
    ok = df[df["状態"] != "欠損"]
    wide = ok.pivot_table(index="日時", columns="ステーションID", values="貸出可能台数", aggfunc="last")
    order = info.sort_values(["区", "ステーションID"])["ステーションID"].tolist()
    order = [s for s in order if s in wide.columns]
    wide = wide[order].sort_index(ascending=False)

    # 見出し3行：区，ステーション名，容量
    heads = [("区", "区"), ("ステーション名", "ステーション名"), ("容量", "容量")]
    for k, (label, key) in enumerate(heads):
        r = 4 + k
        c = ws.cell(row=r, column=1, value=label)
        c.font, c.fill, c.alignment, c.border = F_HEAD, FILL_HEAD, CENTER, BORDER
        for j, sid in enumerate(order, 2):
            v = info.set_index("ステーションID").loc[sid, key]
            c = ws.cell(row=r, column=j, value=v)
            c.font, c.fill, c.alignment, c.border = F_HEAD, FILL_HEAD, CENTER, BORDER
    ws.row_dimensions[5].height = 75
    ws.column_dimensions["A"].width = 17
    for j in range(2, len(order) + 2):
        ws.column_dimensions[get_column_letter(j)].width = 9

    for i, (t, row) in enumerate(wide.iterrows(), 7):
        c = ws.cell(row=i, column=1, value=t)
        c.font, c.border, c.alignment = F_BODY, BORDER, CENTER
        for j, v in enumerate(row, 2):
            c = ws.cell(row=i, column=j, value=None if pd.isna(v) else int(v))
            c.font, c.border, c.alignment = F_BODY, BORDER, CENTER
    end = 6 + len(wide)
    last_col = get_column_letter(len(order) + 1)
    rng = f"B7:{last_col}{end}"
    # 0台（空車）は青，1台は白，台数が多いほど緑
    ws.conditional_formatting.add(rng, ColorScaleRule(
        start_type="num", start_value=0, start_color="9DC3E6",
        mid_type="num", mid_value=1, mid_color="FFFFFF",
        end_type="max", end_color="63BE7B"))
    # 日付が変わるところに太線
    prev = None
    for i in range(7, end + 1):
        d = str(ws.cell(row=i, column=1).value)[:10]
        if prev is not None and d != prev:
            for j in range(1, len(order) + 2):
                c = ws.cell(row=i, column=j)
                b = c.border
                c.border = Border(left=b.left, right=b.right, top=MED, bottom=b.bottom)
        prev = d
    outline(ws, 4, 1, end, len(order) + 1)
    ws.freeze_panes = "B7"


def sheet_all(wb, df):
    ws = wb.create_sheet(f"直近{RECENT_DAYS}日の全データ")
    title(ws, f"15分ごとの全記録（直近{RECENT_DAYS}日）",
          "見出しの▼で絞り込めます．全期間のデータは summary/all_data.csv にあります．")
    t = df.sort_values(["日時", "区", "ステーションID"], ascending=[False, True, True])
    t = t.drop(columns=["日付", "時刻"])
    end = write_table(ws, t, 4,
                      {"日時": 17, "曜日": 6, "区": 6, "ステーションID": 13, "ステーション名": 44,
                       "容量": 7, "貸出可能台数": 12, "空き枠": 8, "貸出不可台数": 12, "状態": 8,
                       "ステーション最終更新": 20},
                      text_cols={"日時", "曜日", "区", "ステーションID", "ステーション名", "状態", "ステーション最終更新"})
    status_col = get_column_letter(list(t.columns).index("状態") + 1)
    last_col = get_column_letter(len(t.columns))
    rng = f"A5:{last_col}{end}"
    for name, fill in STATUS_FILL.items():
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'${status_col}5="{name}"'], fill=fill))
    outline(ws, 4, 1, end, len(t.columns))
    ws.auto_filter.ref = f"A4:{last_col}{end}"
    ws.freeze_panes = "A5"


def sheet_help(wb):
    ws = wb.create_sheet("説明")
    title(ws, "このファイルの見方", "毎日 0:20 ごろ（日本時間）に自動で作り直されます．")
    rows = [
        ("項目", "意味"),
        ("容量", "ステーションのラック数（物理的に止められる最大台数）"),
        ("貸出可能台数", "今すぐ借りられる自転車の台数"),
        ("空き枠", "今すぐ返却できる枠の数．0なら満車で返却できない"),
        ("貸出不可台数", "容量 −（貸出可能台数＋空き枠）．故障・電池切れ・予約中（最大30分）などで貸せない車両が枠を占めている数の目安"),
        ("満車（赤）", "空き枠が0．返却したい人が返せない状態"),
        ("空車（青）", "貸出可能台数が0．借りたい人が借りられない状態"),
        ("休止（黄）", "ステーションが貸出を停止している"),
        ("欠損（灰）", "その回にAPIからステーションの情報が得られなかった"),
        ("データの出典", "公共交通オープンデータ協議会（ODPT）が公開する HELLO CYCLING の GBFS（station_status / station_information）を15分ごとに取得"),
    ]
    for i, (a, b) in enumerate(rows, 4):
        ca, cb = ws.cell(row=i, column=1, value=a), ws.cell(row=i, column=2, value=b)
        for c in (ca, cb):
            c.font, c.border = F_BODY, BORDER
            c.alignment = Alignment(vertical="center", wrap_text=True)
        if i == 4:
            for c in (ca, cb):
                c.font, c.fill = F_HEAD, FILL_HEAD
        key = a.split("（")[0]
        if key in STATUS_FILL:
            ca.fill = STATUS_FILL[key]
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 90
    outline(ws, 4, 1, 3 + len(rows), 2)


def main():
    df = pd.read_csv(f"{OUT_DIR}/all_data.csv", dtype={"ステーションID": str}, encoding="utf-8-sig")
    daily = pd.read_csv(f"{OUT_DIR}/daily_by_station.csv", dtype={"ステーションID": str}, encoding="utf-8-sig")

    info = (df.dropna(subset=["ステーション名"])
              .groupby("ステーションID")[["区", "ステーション名", "容量"]].last().reset_index())

    times = pd.to_datetime(df["日時"], format="%Y/%m/%d %H:%M")
    recent = df[times >= times.max() - timedelta(days=RECENT_DAYS)]

    wb = Workbook()
    sheet_latest(wb, df)
    sheet_daily(wb, daily)
    sheet_wide(wb, recent, info)
    sheet_all(wb, recent)
    sheet_help(wb)
    wb.save(f"{OUT_DIR}/summary.xlsx")
    print(f"✅ Excelを作成しました → {OUT_DIR}/summary.xlsx")


if __name__ == "__main__":
    main()
