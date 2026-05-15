import streamlit as st
import pandas as pd
from datetime import date, datetime
from pathlib import Path
import uuid
import io
from urllib.parse import quote_plus

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase import pdfmetrics


# =========================================================
# にゃんとも相談管理システム Ver1.6（案件選択改善版）
# ---------------------------------------------------------
# 追加機能：
# ・PDF出力
# ・AI要約用メモ作成
# ・GoogleMapリンク
# ・写真管理
#
# 方針：
# ・診断しない
# ・結論を急がせない
# ・未確定と保留を分けて記録する
# ・今はExcel保存で安定運用、将来SQLite化しやすい構造
# =========================================================

st.set_page_config(
    page_title="にゃんとも相談管理システム",
    page_icon="🐾",
    layout="wide"
)

DATA_FILE = Path("nyantomo_consultation_data.xlsx")
PHOTO_DIR = Path("photos")
PHOTO_DIR.mkdir(exist_ok=True)

CLIENT_COLUMNS = [
    "client_id", "登録日時", "お名前", "年代", "地域", "連絡方法", "相談者の立場", "備考"
]

CASE_COLUMNS = [
    "case_id", "client_id", "登録日時", "相談日", "案件名", "案件種別", "現在ステータス",
    "今いちばん近い状態", "住まいの状態", "猫との関係", "家族との温度差", "急がされている感じ",
    "気になること", "今は決めたくないこと", "まず確認したいこと", "自由メモ", "内部メモ", "次回確認すること"
]

HISTORY_COLUMNS = [
    "history_id", "case_id", "client_id", "記録日時", "記録日", "記録種別",
    "状態変更前", "状態変更後", "相談記録", "次回アクション", "内部メモ"
]

PROPERTY_COLUMNS = [
    "property_id", "case_id", "client_id", "登録日時", "物件名", "所在地", "物件状態",
    "空き家状態", "鍵預かり", "近隣不安", "管理頻度", "メモ"
]

CAT_COLUMNS = [
    "cat_id", "case_id", "client_id", "登録日時", "猫の名前", "年齢", "頭数",
    "現在の暮らし", "気になること", "預け先候補", "メモ"
]

FAMILY_COLUMNS = [
    "family_id", "case_id", "client_id", "登録日時", "関係者名", "続柄",
    "連絡可否", "温度感", "関係メモ"
]

PHOTO_COLUMNS = [
    "photo_id", "case_id", "client_id", "登録日時", "写真種別", "ファイル名", "保存先", "説明"
]

STATUS_ORDER = [
    "未対応", "初回相談前", "初回相談済", "情報整理中", "保留中",
    "見守り中", "継続相談", "専門家紹介済", "終了"
]


def make_id(prefix: str) -> str:
    return f"{prefix}_{uuid.uuid4().hex[:10]}"


def blank_df(columns):
    return pd.DataFrame(columns=columns)


def load_sheet(sheet_name, columns):
    if not DATA_FILE.exists():
        return blank_df(columns)
    try:
        df = pd.read_excel(DATA_FILE, sheet_name=sheet_name, dtype=str)
        for col in columns:
            if col not in df.columns:
                df[col] = ""
        return df[columns].fillna("")
    except Exception:
        return blank_df(columns)


def save_all(dataframes):
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as writer:
        for sheet_name, df in dataframes.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)


def load_all():
    return {
        "clients": load_sheet("clients", CLIENT_COLUMNS),
        "cases": load_sheet("cases", CASE_COLUMNS),
        "history": load_sheet("history", HISTORY_COLUMNS),
        "properties": load_sheet("properties", PROPERTY_COLUMNS),
        "cats": load_sheet("cats", CAT_COLUMNS),
        "family": load_sheet("family", FAMILY_COLUMNS),
        "photos": load_sheet("photos", PHOTO_COLUMNS),
    }


def get_client_label(row):
    return f"{row.get('お名前','')}｜{row.get('地域','')}｜{row.get('client_id','')}"


def get_case_label(row):
    """案件選択用ラベル。
    視認性を上げるため、相談者名を先頭に表示する。
    selected_id_from_label() が使えるよう、case_id は必ず最後に置く。
    """
    client_id = str(row.get("client_id", ""))
    client_name = "相談者未登録"

    try:
        clients_df = data.get("clients", pd.DataFrame())
        if not clients_df.empty and "client_id" in clients_df.columns:
            matched = clients_df[clients_df["client_id"] == client_id]
            if not matched.empty:
                name = str(matched.iloc[0].get("お名前", "")).strip()
                if name:
                    client_name = name
    except Exception:
        pass

    consult_date = str(row.get("相談日", "")).strip()
    case_title = str(row.get("案件名", "")).strip() or "案件名未入力"
    status = str(row.get("現在ステータス", "")).strip() or "ステータス未設定"
    case_id = str(row.get("case_id", "")).strip()

    if consult_date:
        return f"{client_name}｜{case_title}｜{status}｜{consult_date}｜{case_id}"
    return f"{client_name}｜{case_title}｜{status}｜{case_id}"


def selected_id_from_label(label):
    if not label or "｜" not in label:
        return ""
    return label.split("｜")[-1]




def get_client_name_by_id(data, client_id):
    clients_df = data.get("clients", pd.DataFrame())
    if clients_df.empty or "client_id" not in clients_df.columns:
        return ""
    matched = clients_df[clients_df["client_id"].astype(str) == str(client_id)]
    if matched.empty:
        return ""
    return str(matched.iloc[0].get("お名前", "")).strip()


def find_duplicate_active_cases_by_client_name(data, client_id):
    """同じ相談者名の未終了案件がある場合、案件の二重登録を防ぐ。"""
    client_name = get_client_name_by_id(data, client_id)
    if not client_name:
        return pd.DataFrame()

    clients_df = data.get("clients", pd.DataFrame())
    cases_df = data.get("cases", pd.DataFrame())
    if clients_df.empty or cases_df.empty:
        return pd.DataFrame()

    same_name_clients = clients_df[clients_df["お名前"].astype(str).str.strip() == client_name]
    same_client_ids = same_name_clients["client_id"].astype(str).tolist()
    dup = cases_df[
        (cases_df["client_id"].astype(str).isin(same_client_ids))
        & (cases_df["現在ステータス"].astype(str) != "終了")
    ].copy()

    if dup.empty:
        return dup

    dup = dup.merge(
        clients_df[["client_id", "お名前", "地域"]],
        on="client_id",
        how="left"
    )
    return dup


def build_case_memo(data, case_id):
    case_df = data["cases"][data["cases"]["case_id"] == case_id]
    if case_df.empty:
        return ""
    case_row = case_df.iloc[0]

    client_df = data["clients"][data["clients"]["client_id"] == case_row["client_id"]]
    client_row = client_df.iloc[0] if not client_df.empty else pd.Series(dtype=str)

    history_df = data["history"][data["history"]["case_id"] == case_id]
    prop_df = data["properties"][data["properties"]["case_id"] == case_id]
    cat_df = data["cats"][data["cats"]["case_id"] == case_id]
    fam_df = data["family"][data["family"]["case_id"] == case_id]
    photo_df = data["photos"][data["photos"]["case_id"] == case_id]

    history_text = "\n".join([
        f"- {h['記録日']}｜{h['記録種別']}｜{h['状態変更前']} → {h['状態変更後']}｜{h['相談記録']}"
        for _, h in history_df.iterrows()
    ]) or "未登録"

    prop_text = "\n".join([
        f"- {p['物件名']}｜{p['所在地']}｜{p['物件状態']}｜{p['空き家状態']}｜管理頻度：{p['管理頻度']}"
        for _, p in prop_df.iterrows()
    ]) or "未登録"

    cat_text = "\n".join([
        f"- {c['猫の名前']}｜年齢：{c['年齢']}｜頭数：{c['頭数']}｜{c['現在の暮らし']}｜気になること：{c['気になること']}"
        for _, c in cat_df.iterrows()
    ]) or "未登録"

    fam_text = "\n".join([
        f"- {f['関係者名']}｜{f['続柄']}｜{f['温度感']}｜{f['関係メモ']}"
        for _, f in fam_df.iterrows()
    ]) or "未登録"

    photo_text = "\n".join([
        f"- {p['写真種別']}｜{p['ファイル名']}｜{p['説明']}"
        for _, p in photo_df.iterrows()
    ]) or "未登録"

    return f"""【にゃんとも相談整理メモ】

作成日：{date.today().strftime('%Y-%m-%d')}

■ 相談者
お名前：{client_row.get('お名前', '')}
地域：{client_row.get('地域', '')}
年代：{client_row.get('年代', '')}
連絡方法：{client_row.get('連絡方法', '')}
相談者の立場：{client_row.get('相談者の立場', '')}

■ 案件
案件名：{case_row['案件名']}
案件種別：{case_row['案件種別']}
現在ステータス：{case_row['現在ステータス']}
相談日：{case_row['相談日']}

■ 今いちばん近い状態
{case_row['今いちばん近い状態']}

■ 住まいの状態
{case_row['住まいの状態']}

■ 猫との関係
{case_row['猫との関係']}

■ 気になること
{case_row['気になること']}

■ 家族との温度差
{case_row['家族との温度差']}

■ 急がされている感じ
{case_row['急がされている感じ']}

■ 今は決めたくないこと
{case_row['今は決めたくないこと']}

■ まず確認したいこと
{case_row['まず確認したいこと']}

■ 空き家カード
{prop_text}

■ 猫情報カード
{cat_text}

■ 家族関係メモ
{fam_text}

■ 写真記録
{photo_text}

■ 相談履歴
{history_text}

■ 内部メモ
{case_row['内部メモ']}

■ 次回確認すること
{case_row['次回確認すること']}

※このメモは、判断を急がせず、状況を整理するための内部記録です。
※法的判断・医療判断・不動産判断を断定するものではありません。
"""


def build_ai_summary(data, case_id):
    case_df = data["cases"][data["cases"]["case_id"] == case_id]
    if case_df.empty:
        return ""
    c = case_df.iloc[0]
    return f"""【AI要約・下書き】

■ 現在の状態
この案件は「{c['現在ステータス']}」の状態です。
相談者は「{c['今いちばん近い状態']}」に近く、住まいについては「{c['住まいの状態']}」という状況です。

■ 猫との関係
{c['猫との関係']}

■ 不安・気になること
{c['気になること']}

■ 家族との温度差
{c['家族との温度差']}

■ 急がされている感じ
{c['急がされている感じ']}

■ いま決めないでよいこと
{c['今は決めたくないこと']}

■ 次回確認すること
{c['次回確認すること']}

■ 注意
この要約は判断を代行するものではありません。
未確定の内容は未確定のまま扱い、相談者に結論を急がせない前提で確認します。
"""


def build_ai_prompt(memo):
    return f"""あなたは、にゃんとも相談管理システムの記録整理係です。
以下の相談メモをもとに、判断を急がせない内部要約を作成してください。

【重要ルール】
・法的判断、医療判断、不動産判断を断定しない。
・「売るべき」「貸すべき」「施設に入るべき」などの結論を出さない。
・事実、未確定、保留、次回確認事項を分ける。
・相談者を責めない。
・家族間の温度差を対立として煽らない。
・猫、住まい、人の暮らしを分断せずに整理する。
・最後に「次回確認すること」を3つ以内で出す。

【相談メモ】
{memo}
"""


def make_pdf_bytes(text):
    buffer = io.BytesIO()
    pdfmetrics.registerFont(UnicodeCIDFont("HeiseiMin-W3"))

    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    x = 40
    y = height - 45
    line_height = 14
    max_chars = 45

    c.setFont("HeiseiMin-W3", 13)
    c.drawString(x, y, "にゃんとも相談整理メモ")
    y -= 25

    c.setFont("HeiseiMin-W3", 9)

    for raw_line in text.splitlines():
        line = raw_line.rstrip()
        if not line:
            y -= line_height
        else:
            chunks = [line[i:i + max_chars] for i in range(0, len(line), max_chars)]
            for chunk in chunks:
                if y < 45:
                    c.showPage()
                    c.setFont("HeiseiMin-W3", 9)
                    y = height - 45
                c.drawString(x, y, chunk)
                y -= line_height

    c.save()
    buffer.seek(0)
    return buffer.getvalue()



# ---------------------------------------------------------
# Ver1.3 追加：全データ共通の検索・更新・削除
# ---------------------------------------------------------
TABLE_CONFIG = {
    "相談者データ": {
        "key": "clients",
        "id_col": "client_id",
        "columns": CLIENT_COLUMNS,
        "protected_cols": ["client_id", "登録日時"],
    },
    "案件データ": {
        "key": "cases",
        "id_col": "case_id",
        "columns": CASE_COLUMNS,
        "protected_cols": ["case_id", "client_id", "登録日時"],
    },
    "相談履歴データ": {
        "key": "history",
        "id_col": "history_id",
        "columns": HISTORY_COLUMNS,
        "protected_cols": ["history_id", "case_id", "client_id", "記録日時"],
    },
    "空き家カード": {
        "key": "properties",
        "id_col": "property_id",
        "columns": PROPERTY_COLUMNS,
        "protected_cols": ["property_id", "case_id", "client_id", "登録日時"],
    },
    "猫情報カード": {
        "key": "cats",
        "id_col": "cat_id",
        "columns": CAT_COLUMNS,
        "protected_cols": ["cat_id", "case_id", "client_id", "登録日時"],
    },
    "家族関係メモ": {
        "key": "family",
        "id_col": "family_id",
        "columns": FAMILY_COLUMNS,
        "protected_cols": ["family_id", "case_id", "client_id", "登録日時"],
    },
    "写真データ": {
        "key": "photos",
        "id_col": "photo_id",
        "columns": PHOTO_COLUMNS,
        "protected_cols": ["photo_id", "case_id", "client_id", "登録日時", "保存先"],
    },
}


def filter_df_keyword(df, keyword):
    if df.empty or not keyword:
        return df
    return df[df.astype(str).apply(lambda row: row.str.contains(keyword, case=False, na=False).any(), axis=1)]


def sync_case_status_from_history(data, case_id):
    history_df = data["history"][data["history"]["case_id"] == case_id]
    if history_df.empty:
        return data
    latest = history_df.iloc[-1]
    new_status = latest.get("状態変更後", "")
    if new_status:
        data["cases"].loc[data["cases"]["case_id"] == case_id, "現在ステータス"] = new_status
    return data


def delete_related_records(data, table_key, id_col, record_id):
    """親データ削除時は関連データも一緒に削除する。"""
    if table_key == "clients":
        client_id = record_id
        photo_rows = data["photos"][data["photos"]["client_id"] == client_id]
        for _, p in photo_rows.iterrows():
            path = Path(str(p.get("保存先", "")))
            if path.exists():
                try:
                    path.unlink()
                except Exception:
                    pass

        for key in ["clients", "cases", "history", "properties", "cats", "family", "photos"]:
            data[key] = data[key][data[key]["client_id"] != client_id] if "client_id" in data[key].columns else data[key]

    elif table_key == "cases":
        case_id = record_id
        photo_rows = data["photos"][data["photos"]["case_id"] == case_id]
        for _, p in photo_rows.iterrows():
            path = Path(str(p.get("保存先", "")))
            if path.exists():
                try:
                    path.unlink()
                except Exception:
                    pass

        for key in ["cases", "history", "properties", "cats", "family", "photos"]:
            data[key] = data[key][data[key]["case_id"] != case_id] if "case_id" in data[key].columns else data[key]

    elif table_key == "photos":
        photo_rows = data["photos"][data["photos"][id_col] == record_id]
        for _, p in photo_rows.iterrows():
            path = Path(str(p.get("保存先", "")))
            if path.exists():
                try:
                    path.unlink()
                except Exception:
                    pass
        data["photos"] = data["photos"][data["photos"][id_col] != record_id]

    else:
        data[table_key] = data[table_key][data[table_key][id_col] != record_id]

    return data


def render_search_update_delete(data):
    st.subheader("🔎 全データ検索・更新・削除")
    st.caption("相談者、案件、相談履歴、空き家、猫、家族、写真の各データを検索・編集・削除できます。")

    table_name = st.selectbox("対象データを選択", list(TABLE_CONFIG.keys()), key="crud_table_select")
    cfg = TABLE_CONFIG[table_name]
    table_key = cfg["key"]
    id_col = cfg["id_col"]
    columns = cfg["columns"]
    protected_cols = cfg["protected_cols"]

    df = data[table_key].copy()
    if df.empty:
        st.info("このデータはまだ登録されていません。")
        return data

    keyword = st.text_input("検索キーワード", placeholder="名前・地域・住所・猫の名前・メモなどで検索", key=f"crud_keyword_{table_key}")
    filtered_df = filter_df_keyword(df, keyword)

    st.write(f"検索結果：{len(filtered_df)}件")
    st.dataframe(filtered_df, use_container_width=True)

    st.divider()
    st.markdown("### 1件ずつ更新")

    if filtered_df.empty:
        st.info("検索結果がありません。")
    else:
        label_options = []
        for _, row in filtered_df.iterrows():
            rid = row[id_col]
            title_parts = []
            for col in columns:
                value = str(row.get(col, ""))
                if value and col != id_col:
                    title_parts.append(value)
                if len(title_parts) >= 2:
                    break
            label_options.append(f"{rid}｜{'｜'.join(title_parts)}")

        selected_label = st.selectbox("編集するデータを選択", label_options, key=f"crud_edit_select_{table_key}")
        selected_id = selected_label.split("｜")[0]
        row_df = df[df[id_col] == selected_id]

        if not row_df.empty:
            row = row_df.iloc[0]
            with st.form(f"edit_form_{table_key}"):
                new_values = {}
                for col in columns:
                    value = str(row.get(col, ""))
                    if col in protected_cols:
                        st.text_input(col, value=value, disabled=True)
                        new_values[col] = value
                    else:
                        if len(value) > 35 or "メモ" in col or "こと" in col or "記録" in col or "説明" in col:
                            new_values[col] = st.text_area(col, value=value, height=100)
                        else:
                            new_values[col] = st.text_input(col, value=value)

                submitted = st.form_submit_button("この内容で更新する")
                if submitted:
                    for col, value in new_values.items():
                        data[table_key].loc[data[table_key][id_col] == selected_id, col] = value

                    if table_key == "history":
                        case_id = str(row.get("case_id", ""))
                        if case_id:
                            data = sync_case_status_from_history(data, case_id)

                    save_all(data)
                    st.success("更新しました。")
                    st.rerun()

    st.divider()
    st.markdown("### 削除")
    st.warning("削除は元に戻せません。削除前に、データ管理からExcelをダウンロードしてバックアップしてください。")

    if not filtered_df.empty:
        delete_options = []
        for _, row in filtered_df.iterrows():
            rid = row[id_col]
            title_parts = []
            for col in columns:
                value = str(row.get(col, ""))
                if value and col != id_col:
                    title_parts.append(value)
                if len(title_parts) >= 2:
                    break
            delete_options.append(f"{rid}｜{'｜'.join(title_parts)}")

        delete_label = st.selectbox("削除するデータを選択", delete_options, key=f"delete_{table_key}")
        delete_id = delete_label.split("｜")[0]

        if table_key == "clients":
            st.error("相談者を削除すると、その相談者に紐づく案件・履歴・空き家・猫・家族・写真データも削除されます。")
        elif table_key == "cases":
            st.error("案件を削除すると、その案件に紐づく履歴・空き家・猫・家族・写真データも削除されます。")

        confirm = st.checkbox("本当に削除します", key=f"confirm_delete_{table_key}")
        if st.button("削除する", type="primary", disabled=not confirm):
            data = delete_related_records(data, table_key, id_col, delete_id)
            save_all(data)
            st.success("削除しました。")
            st.rerun()

    return data



# ---------------------------------------------------------
# Ver1.6 追加：業務改善ダッシュボード・次アクション・入力不足チェック
# ---------------------------------------------------------

def parse_date_safe(value):
    """YYYY-MM-DDや日時文字列をdateへ安全に変換する。"""
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    for fmt in ["%Y-%m-%d", "%Y-%m-%d %H:%M:%S"]:
        try:
            return datetime.strptime(text[:19], fmt).date()
        except Exception:
            pass
    try:
        return pd.to_datetime(text).date()
    except Exception:
        return None


def latest_case_activity_date(data, case_id):
    dates = []
    for key, date_col in [
        ("history", "記録日"),
        ("properties", "登録日時"),
        ("cats", "登録日時"),
        ("family", "登録日時"),
        ("photos", "登録日時"),
    ]:
        df = data[key]
        if df.empty or "case_id" not in df.columns:
            continue
        rows = df[df["case_id"] == case_id]
        if date_col not in rows.columns:
            continue
        for v in rows[date_col].tolist():
            d = parse_date_safe(v)
            if d:
                dates.append(d)
    return max(dates) if dates else None


def case_silent_status(days):
    if days is None:
        return "記録なし"
    if days >= 90:
        return "90日以上未更新"
    if days >= 60:
        return "60日以上未更新"
    if days >= 30:
        return "30日以上未更新"
    return "確認中"


def build_case_timeline(data, case_id):
    rows = []

    history_df = data["history"][data["history"]["case_id"] == case_id]
    for _, h in history_df.iterrows():
        rows.append({
            "日付": h.get("記録日", ""),
            "種別": h.get("記録種別", "履歴"),
            "内容": h.get("相談記録", ""),
            "状態": f"{h.get('状態変更前','')} → {h.get('状態変更後','')}",
            "次回": h.get("次回アクション", ""),
        })

    photo_df = data["photos"][data["photos"]["case_id"] == case_id]
    for _, p in photo_df.iterrows():
        rows.append({
            "日付": str(p.get("登録日時", ""))[:10],
            "種別": f"写真：{p.get('写真種別','')}",
            "内容": p.get("説明", ""),
            "状態": "",
            "次回": "",
        })

    prop_df = data["properties"][data["properties"]["case_id"] == case_id]
    for _, p in prop_df.iterrows():
        rows.append({
            "日付": str(p.get("登録日時", ""))[:10],
            "種別": "空き家カード",
            "内容": f"{p.get('物件名','')}｜{p.get('所在地','')}｜{p.get('物件状態','')}｜{p.get('空き家状態','')}",
            "状態": p.get("管理頻度", ""),
            "次回": "",
        })

    cat_df = data["cats"][data["cats"]["case_id"] == case_id]
    for _, c in cat_df.iterrows():
        rows.append({
            "日付": str(c.get("登録日時", ""))[:10],
            "種別": "猫情報",
            "内容": f"{c.get('猫の名前','')}｜{c.get('現在の暮らし','')}｜{c.get('気になること','')}",
            "状態": "",
            "次回": "",
        })

    fam_df = data["family"][data["family"]["case_id"] == case_id]
    for _, f in fam_df.iterrows():
        rows.append({
            "日付": str(f.get("登録日時", ""))[:10],
            "種別": "家族関係",
            "内容": f"{f.get('関係者名','')}｜{f.get('続柄','')}｜{f.get('温度感','')}｜{f.get('関係メモ','')}",
            "状態": "",
            "次回": "",
        })

    timeline = pd.DataFrame(rows, columns=["日付", "種別", "内容", "状態", "次回"])
    if timeline.empty:
        return timeline
    timeline["_sort"] = pd.to_datetime(timeline["日付"], errors="coerce")
    timeline = timeline.sort_values("_sort", ascending=False).drop(columns=["_sort"])
    return timeline.fillna("")


def build_case_home_dataframe(data):
    """案件ホーム用の一覧を作る。Excel保存のままでも業務判断しやすいように集約する。"""
    rows = []
    today = date.today()
    for _, c in data["cases"].iterrows():
        case_id = c.get("case_id", "")
        client_id = c.get("client_id", "")
        client_rows = data["clients"][data["clients"]["client_id"] == client_id]
        client_name = client_rows.iloc[0].get("お名前", "") if not client_rows.empty else ""
        latest = latest_case_activity_date(data, case_id)
        days = (today - latest).days if latest else None

        prop_count = len(data["properties"][data["properties"]["case_id"] == case_id]) if not data["properties"].empty else 0
        cat_count = len(data["cats"][data["cats"]["case_id"] == case_id]) if not data["cats"].empty else 0
        family_count = len(data["family"][data["family"]["case_id"] == case_id]) if not data["family"].empty else 0
        history_count = len(data["history"][data["history"]["case_id"] == case_id]) if not data["history"].empty else 0

        missing = []
        if not str(c.get("次回確認すること", "")).strip():
            missing.append("次回確認")
        if prop_count == 0 and c.get("住まいの状態", "") in ["空き家になっている", "近いうちに空き家になりそう", "相続後そのまま", "売却・賃貸を迷っている"]:
            missing.append("空き家カード")
        if cat_count == 0 and c.get("猫との関係", "") not in ["未選択", "猫はいない", ""]:
            missing.append("猫情報")
        if family_count == 0 and c.get("家族との温度差", "") in ["少しある", "かなりある", "まだ話せていない"]:
            missing.append("家族メモ")

        status = c.get("現在ステータス", "")
        if status == "終了":
            next_action_label = "終了済"
        elif days is None:
            next_action_label = "まず初回記録を確認"
        elif days >= 90:
            next_action_label = "継続要否を確認"
        elif days >= 60:
            next_action_label = "近況確認の候補"
        elif days >= 30:
            next_action_label = "静かに確認"
        elif str(c.get("次回確認すること", "")).strip():
            next_action_label = "次回確認あり"
        else:
            next_action_label = "保留継続"

        if status == "終了":
            priority = "完了"
        elif days is not None and days >= 90:
            priority = "高"
        elif days is not None and days >= 60:
            priority = "中"
        elif missing:
            priority = "確認"
        else:
            priority = "通常"

        rows.append({
            "優先": priority,
            "次アクション": next_action_label,
            "案件名": c.get("案件名", ""),
            "相談者": client_name,
            "現在ステータス": status,
            "案件種別": c.get("案件種別", ""),
            "最終更新": latest.strftime("%Y-%m-%d") if latest else "",
            "未更新日数": days if days is not None else "",
            "静かな確認": case_silent_status(days),
            "入力不足": "、".join(missing),
            "次回確認すること": c.get("次回確認すること", ""),
            "履歴数": history_count,
            "空き家": prop_count,
            "猫": cat_count,
            "家族": family_count,
            "case_id": case_id,
        })
    return pd.DataFrame(rows)


def render_case_home(data):
    st.subheader("🏡 案件ホーム Ver1.6")
    st.caption("今日見るべき案件、止まっている案件、入力が足りない案件を一画面で確認します。")

    if data["cases"].empty:
        st.info("まだ案件がありません。先に相談者登録・案件登録をしてください。")
        return data

    home_df = build_case_home_dataframe(data)

    active_df = home_df[home_df["現在ステータス"] != "終了"]
    high_df = active_df[active_df["優先"].isin(["高", "中"])]
    missing_df = active_df[active_df["入力不足"].astype(str).str.len() > 0]
    next_df = active_df[active_df["次アクション"].isin(["次回確認あり", "静かに確認", "近況確認の候補", "継続要否を確認", "まず初回記録を確認"])]

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("全案件", len(home_df))
    c2.metric("進行中", len(active_df))
    c3.metric("今日見る候補", len(next_df))
    c4.metric("60日以上未更新", len(high_df))
    c5.metric("入力不足", len(missing_df))

    st.markdown("### 今日見る案件")
    if next_df.empty:
        st.success("今日すぐ確認すべき案件はありません。")
    else:
        st.dataframe(
            next_df[["優先", "次アクション", "案件名", "相談者", "現在ステータス", "未更新日数", "入力不足", "次回確認すること", "case_id"]],
            use_container_width=True,
            hide_index=True,
        )

    with st.expander("入力不足チェック", expanded=not missing_df.empty):
        st.caption("急がせるためではなく、記録の穴をなくして安心して保留するための確認です。")
        if missing_df.empty:
            st.success("大きな入力不足は見つかりません。")
        else:
            st.dataframe(
                missing_df[["案件名", "相談者", "現在ステータス", "入力不足", "空き家", "猫", "家族", "case_id"]],
                use_container_width=True,
                hide_index=True,
            )

    st.divider()
    st.markdown("### 案件一覧")
    col_a, col_b, col_c = st.columns([1, 1, 2])
    with col_a:
        show_active = st.checkbox("終了以外を中心に見る", value=True, key="home_active_only")
    with col_b:
        priority_filter = st.multiselect("優先度", ["高", "中", "確認", "通常", "完了"], default=[], key="home_priority_filter")
    with col_c:
        keyword = st.text_input("案件ホーム検索", placeholder="案件名・相談者・次回確認など", key="home_keyword")

    view_df = home_df.copy()
    if show_active:
        view_df = view_df[view_df["現在ステータス"] != "終了"]
    if priority_filter:
        view_df = view_df[view_df["優先"].isin(priority_filter)]

    status_filter = st.multiselect("ステータス絞り込み", STATUS_ORDER, default=[], key="home_status_filter")
    if status_filter:
        view_df = view_df[view_df["現在ステータス"].isin(status_filter)]
    view_df = filter_df_keyword(view_df, keyword)

    st.dataframe(view_df, use_container_width=True, hide_index=True)

    st.markdown("### 業務改善メモ")
    st.info("Ver1.6では、案件を“探す”時間を減らし、今日見る案件・入力不足・止まっている案件を先に出す設計にしています。")

    return data

def render_case_dashboard(data):
    st.subheader("🗂 案件ダッシュボード")
    st.caption("case_idを中心に、相談者・履歴・空き家・猫・家族・写真を一画面で確認します。ゴールはステータスを『終了』まで進めることです。")

    if data["cases"].empty:
        st.info("先に案件を登録してください。")
        return data

    active_only = st.checkbox("終了以外の案件だけ表示", value=True, key="casehub_active_only")
    case_df = data["cases"].copy()
    if active_only:
        case_df = case_df[case_df["現在ステータス"] != "終了"]

    if case_df.empty:
        st.success("表示対象の案件はありません。")
        return data

    status_filter = st.multiselect(
        "ステータスで絞り込み",
        STATUS_ORDER,
        default=[],
        key="casehub_status_filter"
    )
    if status_filter:
        case_df = case_df[case_df["現在ステータス"].isin(status_filter)]

    keyword = st.text_input("案件検索", placeholder="案件名・気になること・メモなど", key="casehub_keyword")
    case_df = filter_df_keyword(case_df, keyword)

    if case_df.empty:
        st.info("条件に合う案件がありません。")
        return data

    case_labels = [get_case_label(row) for _, row in case_df.iterrows()]
    selected_case_label = st.selectbox("案件を選択", case_labels, key="casehub_select_case")
    case_id = selected_id_from_label(selected_case_label)

    case_rows = data["cases"][data["cases"]["case_id"] == case_id]
    if case_rows.empty:
        st.error("案件が見つかりません。")
        return data

    case_row = case_rows.iloc[0]
    client_id = case_row["client_id"]
    client_rows = data["clients"][data["clients"]["client_id"] == client_id]
    client_row = client_rows.iloc[0] if not client_rows.empty else pd.Series(dtype=str)

    status = str(case_row.get("現在ステータス", ""))
    try:
        progress = int((STATUS_ORDER.index(status) + 1) / len(STATUS_ORDER) * 100)
    except ValueError:
        progress = 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("現在ステータス", status)
    col2.metric("相談者", str(client_row.get("お名前", "未登録")))
    col3.metric("案件種別", str(case_row.get("案件種別", "")))
    col4.metric("終了までの目安", f"{progress}%")
    st.progress(progress)

    with st.container(border=True):
        st.markdown("### 案件の現在地")
        c1, c2 = st.columns(2)
        with c1:
            st.write(f"**案件名：** {case_row.get('案件名','')}")
            st.write(f"**相談日：** {case_row.get('相談日','')}")
            st.write(f"**今いちばん近い状態：** {case_row.get('今いちばん近い状態','')}")
            st.write(f"**住まいの状態：** {case_row.get('住まいの状態','')}")
        with c2:
            st.write(f"**猫との関係：** {case_row.get('猫との関係','')}")
            st.write(f"**家族との温度差：** {case_row.get('家族との温度差','')}")
            st.write(f"**急がされている感じ：** {case_row.get('急がされている感じ','')}")
            st.write(f"**次回確認すること：** {case_row.get('次回確認すること','')}")

    st.markdown("### 状態を進める")
    with st.form(f"casehub_status_form_{case_id}"):
        c1, c2 = st.columns(2)
        with c1:
            new_status = st.selectbox(
                "新しいステータス",
                STATUS_ORDER,
                index=STATUS_ORDER.index(status) if status in STATUS_ORDER else 0,
            )
            record_date = st.date_input("記録日", value=date.today(), key=f"casehub_record_date_{case_id}")
        with c2:
            record_type = st.selectbox("記録種別", ["状態変更", "相談", "電話", "LINE", "メール", "面談", "現地確認", "終了確認", "その他"], key=f"casehub_record_type_{case_id}")
            next_action = st.text_area("次回アクション", key=f"casehub_next_action_{case_id}")
        record = st.text_area("相談記録・判断保留の理由・確認した事実", key=f"casehub_record_{case_id}")
        internal = st.text_area("内部メモ", key=f"casehub_internal_{case_id}")
        submitted = st.form_submit_button("履歴を追加してステータス更新")
        if submitted:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            new_history = {
                "history_id": make_id("hist"),
                "case_id": case_id,
                "client_id": client_id,
                "記録日時": now,
                "記録日": record_date.strftime("%Y-%m-%d"),
                "記録種別": record_type,
                "状態変更前": status,
                "状態変更後": new_status,
                "相談記録": record,
                "次回アクション": next_action,
                "内部メモ": internal,
            }
            data["history"] = pd.concat([data["history"], pd.DataFrame([new_history])], ignore_index=True)
            data["cases"].loc[data["cases"]["case_id"] == case_id, "現在ステータス"] = new_status
            data["cases"].loc[data["cases"]["case_id"] == case_id, "次回確認すること"] = next_action
            save_all(data)
            st.success("履歴を追加し、案件ステータスを更新しました。")
            st.rerun()

    st.divider()
    latest = latest_case_activity_date(data, case_id)
    days = (date.today() - latest).days if latest else None
    st.markdown("### 静かな確認")
    if days is None:
        st.info("この案件は、まだ履歴や関連データが少ない状態です。")
    elif days >= 90:
        st.warning(f"最終更新から{days}日です。終了・継続・保留理由の確認をおすすめします。")
    elif days >= 60:
        st.info(f"最終更新から{days}日です。必要なら一度、状況確認を入れてください。")
    elif days >= 30:
        st.info(f"最終更新から{days}日です。保留理由が変わっていないか確認できます。")
    else:
        st.success(f"最終更新から{days}日です。")

    st.markdown("### タイムライン")
    timeline = build_case_timeline(data, case_id)
    if timeline.empty:
        st.info("まだタイムラインに表示できる記録がありません。")
    else:
        st.dataframe(timeline, use_container_width=True)

    st.divider()
    st.markdown("### この案件に紐づくデータ")
    rel_tabs = st.tabs(["相談者", "相談履歴", "空き家", "猫", "家族", "写真", "AI/PDF用メモ"])

    with rel_tabs[0]:
        st.dataframe(client_rows, use_container_width=True)

    with rel_tabs[1]:
        h = data["history"][data["history"]["case_id"] == case_id].copy()
        if h.empty:
            st.info("相談履歴はまだありません。")
        else:
            st.dataframe(h, use_container_width=True)

    with rel_tabs[2]:
        p_df = data["properties"][data["properties"]["case_id"] == case_id].copy()
        st.dataframe(p_df, use_container_width=True)
        for _, row in p_df.iterrows():
            address = row.get("所在地", "")
            if address:
                map_url = f"https://www.google.com/maps/search/?api=1&query={quote_plus(address)}"
                st.link_button(f"GoogleMapで開く：{row.get('物件名','物件')}", map_url)

    with rel_tabs[3]:
        st.dataframe(data["cats"][data["cats"]["case_id"] == case_id], use_container_width=True)

    with rel_tabs[4]:
        st.dataframe(data["family"][data["family"]["case_id"] == case_id], use_container_width=True)

    with rel_tabs[5]:
        photo_df = data["photos"][data["photos"]["case_id"] == case_id].copy()
        st.dataframe(photo_df, use_container_width=True)
        for _, p in photo_df.iterrows():
            path = Path(str(p.get("保存先", "")))
            if path.exists():
                with st.container(border=True):
                    st.write(f"{p.get('写真種別','')}：{p.get('ファイル名','')}")
                    st.write(p.get("説明", ""))
                    st.image(str(path), width=350)

    with rel_tabs[6]:
        memo = build_case_memo(data, case_id)
        st.text_area("案件統合メモ", memo, height=500)
        pdf_bytes = make_pdf_bytes(memo)
        st.download_button(
            "この案件のPDFをダウンロード",
            data=pdf_bytes,
            file_name=f"nyantomo_case_{case_id}.pdf",
            mime="application/pdf",
        )

    st.divider()
    st.markdown("### 案件一覧")
    st.dataframe(case_df, use_container_width=True)

    return data

data = load_all()

st.title("🐾 にゃんとも相談管理システム Ver1.6（重複登録防止版）")
st.caption("相談を保留のまま管理する現場OS｜案件起点・次アクション・入力不足チェック対応版")

tabs = st.tabs([
    "🏡 案件ホーム",
    "🧑 相談者登録",
    "📝 案件登録",
    "🗂 案件ダッシュボード",
    "📚 相談履歴",
    "⏸ 保留案件一覧",
    "🏠 空き家カード",
    "🗺 GoogleMap",
    "🐈 猫情報カード",
    "👪 家族関係メモ",
    "📷 写真管理",
    "🤖 AI要約",
    "🧾 PDF出力",
    "🔎 検索・更新・削除",
    "📦 データ管理",
])


with tabs[0]:
    data = render_case_home(data)


with tabs[1]:
    st.subheader("相談者登録")

    with st.form("client_form"):
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("お名前")
            age = st.selectbox("年代", ["未選択", "40代", "50代", "60代", "70代", "80代以上"])
            area = st.text_input("地域")
        with col2:
            contact = st.selectbox("連絡方法", ["未選択", "LINE", "メール", "電話", "対面", "その他"])
            position = st.selectbox("相談者の立場", ["未選択", "本人", "家族", "親族", "空き家所有者", "支援者", "その他"])
            note = st.text_area("備考")

        submitted = st.form_submit_button("相談者を登録")
        if submitted:
            if not name:
                st.error("お名前を入力してください。")
            else:
                new_row = {
                    "client_id": make_id("client"),
                    "登録日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "お名前": name,
                    "年代": age,
                    "地域": area,
                    "連絡方法": contact,
                    "相談者の立場": position,
                    "備考": note,
                }
                data["clients"] = pd.concat([data["clients"], pd.DataFrame([new_row])], ignore_index=True)
                save_all(data)
                st.success("相談者を登録しました。")

    st.divider()
    st.dataframe(data["clients"], use_container_width=True)


with tabs[2]:
    st.subheader("案件登録")

    if data["clients"].empty:
        st.info("先に相談者を登録してください。")
    else:
        client_labels = [get_client_label(row) for _, row in data["clients"].iterrows()]

        with st.form("case_form"):
            selected_client_label = st.selectbox("相談者を選択", client_labels)
            client_id = selected_id_from_label(selected_client_label)

            duplicate_cases_preview = find_duplicate_active_cases_by_client_name(data, client_id)
            if not duplicate_cases_preview.empty:
                st.warning("この相談者名では、すでに終了していない案件が登録されています。二重登録防止のため、新規案件登録はできません。既存案件の『案件ダッシュボード』または『相談履歴』から追記してください。")
                st.dataframe(
                    duplicate_cases_preview[["お名前", "地域", "案件名", "案件種別", "現在ステータス", "相談日", "case_id"]],
                    use_container_width=True
                )

            col1, col2 = st.columns(2)
            with col1:
                consult_date = st.date_input("相談日", value=date.today())
                case_title = st.text_input("案件名", value="住まいと猫の相談")
                case_type = st.selectbox("案件種別", ["初回相談", "空き家管理", "猫と住まい", "相続前整理", "高齢期の住まい", "その他"])
                status = st.selectbox("現在ステータス", STATUS_ORDER, index=1)
                current_state = st.selectbox(
                    "今いちばん近い状態",
                    ["未選択", "まだ何も決まっていない", "少し考え始めている", "家族と話し始めた", "急かされている感じがある", "誰にも相談していない", "すでに困りごとが出ている"],
                )
            with col2:
                house_state = st.selectbox(
                    "住まいの状態",
                    ["未選択", "現在住んでいる", "空き家になっている", "近いうちに空き家になりそう", "相続後そのまま", "売却・賃貸を迷っている", "荷物整理が進んでいない"],
                )
                cat_relation = st.selectbox(
                    "猫との関係",
                    ["未選択", "猫と暮らしている", "家族の猫がいる", "猫を残して入院・施設入所が心配", "これから猫と暮らしたい", "保護猫に関心がある", "猫はいない"],
                )
                family_gap = st.selectbox("家族との温度差", ["未選択", "特にない", "少しある", "かなりある", "まだ話せていない"])
                pressure = st.selectbox("急がされている感じ", ["未選択", "ない", "少しある", "強くある", "自分でも焦っている"])

            worries = st.multiselect(
                "気になること",
                ["空き家管理", "相続", "売却", "賃貸", "猫の住まい", "高齢期の暮らし", "家族との意見の違い", "お金", "近所への不安", "何から考えればよいか分からない"],
            )

            not_decide = st.text_area("今は決めたくないこと")
            first_check = st.text_area("まず確認したいこと")
            free_memo = st.text_area("自由メモ")
            internal_memo = st.text_area("内部メモ")
            next_check = st.text_area("次回確認すること")

            submitted = st.form_submit_button("案件を登録")

            if submitted:
                duplicate_cases = find_duplicate_active_cases_by_client_name(data, client_id)
                if not duplicate_cases.empty:
                    st.error("登録できません。同じ相談者名で、終了していない案件がすでにあります。既存案件に履歴を追加してください。")
                    st.dataframe(
                        duplicate_cases[["お名前", "地域", "案件名", "案件種別", "現在ステータス", "相談日", "case_id"]],
                        use_container_width=True
                    )
                    st.stop()

                case_id = make_id("case")
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                new_case = {
                    "case_id": case_id,
                    "client_id": client_id,
                    "登録日時": now,
                    "相談日": consult_date.strftime("%Y-%m-%d"),
                    "案件名": case_title,
                    "案件種別": case_type,
                    "現在ステータス": status,
                    "今いちばん近い状態": current_state,
                    "住まいの状態": house_state,
                    "猫との関係": cat_relation,
                    "家族との温度差": family_gap,
                    "急がされている感じ": pressure,
                    "気になること": "、".join(worries),
                    "今は決めたくないこと": not_decide,
                    "まず確認したいこと": first_check,
                    "自由メモ": free_memo,
                    "内部メモ": internal_memo,
                    "次回確認すること": next_check,
                }

                new_history = {
                    "history_id": make_id("hist"),
                    "case_id": case_id,
                    "client_id": client_id,
                    "記録日時": now,
                    "記録日": consult_date.strftime("%Y-%m-%d"),
                    "記録種別": "案件登録",
                    "状態変更前": "",
                    "状態変更後": status,
                    "相談記録": free_memo,
                    "次回アクション": next_check,
                    "内部メモ": internal_memo,
                }

                data["cases"] = pd.concat([data["cases"], pd.DataFrame([new_case])], ignore_index=True)
                data["history"] = pd.concat([data["history"], pd.DataFrame([new_history])], ignore_index=True)
                save_all(data)
                st.success("案件を登録しました。")

    st.divider()
    st.dataframe(data["cases"], use_container_width=True)


with tabs[3]:
    data = render_case_dashboard(data)

with tabs[4]:
    st.subheader("相談履歴・状態変更")

    if data["cases"].empty:
        st.info("先に案件を登録してください。")
    else:
        case_labels = [get_case_label(row) for _, row in data["cases"].iterrows()]
        selected_case_label = st.selectbox("案件を選択", case_labels, key="history_select_case")
        case_id = selected_id_from_label(selected_case_label)

        case_rows = data["cases"][data["cases"]["case_id"] == case_id]
        if not case_rows.empty:
            case_row = case_rows.iloc[0]
            current_status = case_row["現在ステータス"]
            client_id = case_row["client_id"]

            st.info(f"現在ステータス：{current_status}")

            with st.form("history_form"):
                col1, col2 = st.columns(2)
                with col1:
                    record_date = st.date_input("記録日", value=date.today(), key=f"history_record_date_{case_id}")
                    record_type = st.selectbox("記録種別", ["相談", "電話", "LINE", "メール", "面談", "現地確認", "状態変更", "その他"], key=f"history_record_type_{case_id}")
                with col2:
                    new_status = st.selectbox("状態変更後", STATUS_ORDER, index=STATUS_ORDER.index(current_status) if current_status in STATUS_ORDER else 0, key=f"history_new_status_{case_id}")

                record = st.text_area("相談記録", key=f"history_record_{case_id}")
                next_action = st.text_area("次回アクション", key=f"history_next_action_{case_id}")
                internal = st.text_area("内部メモ", key=f"history_internal_{case_id}")

                submitted = st.form_submit_button("履歴を追加")

                if submitted:
                    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    new_history = {
                        "history_id": make_id("hist"),
                        "case_id": case_id,
                        "client_id": client_id,
                        "記録日時": now,
                        "記録日": record_date.strftime("%Y-%m-%d"),
                        "記録種別": record_type,
                        "状態変更前": current_status,
                        "状態変更後": new_status,
                        "相談記録": record,
                        "次回アクション": next_action,
                        "内部メモ": internal,
                    }
                    data["history"] = pd.concat([data["history"], pd.DataFrame([new_history])], ignore_index=True)
                    data["cases"].loc[data["cases"]["case_id"] == case_id, "現在ステータス"] = new_status
                    save_all(data)
                    st.success("相談履歴を追加しました。")

        st.divider()
        st.dataframe(data["history"][data["history"]["case_id"] == case_id], use_container_width=True)


with tabs[5]:
    st.subheader("保留案件一覧")
    if data["cases"].empty:
        st.info("案件がありません。")
    else:
        hold_df = data["cases"][data["cases"]["現在ステータス"].isin(["保留中", "情報整理中", "見守り中"])]
        if hold_df.empty:
            st.success("現在、保留・見守り中の案件はありません。")
        else:
            st.dataframe(hold_df, use_container_width=True)


with tabs[6]:
    st.subheader("空き家カード")

    if data["cases"].empty:
        st.info("先に案件を登録してください。")
    else:
        case_labels = [get_case_label(row) for _, row in data["cases"].iterrows()]
        selected_case_label = st.selectbox("案件選択（空き家カード）", case_labels, key="property_select_case")
        case_id = selected_id_from_label(selected_case_label)
        case_row = data["cases"][data["cases"]["case_id"] == case_id].iloc[0]
        client_id = case_row["client_id"]

        with st.form("property_form"):
            col1, col2 = st.columns(2)
            with col1:
                property_name = st.text_input("物件名")
                address = st.text_input("所在地")
                property_status = st.selectbox("物件状態", ["未確認", "居住中", "空き家", "一部使用", "売却検討", "賃貸検討", "その他"])
                vacant_status = st.selectbox("空き家状態", ["未確認", "問題なし", "定期確認必要", "劣化あり", "近隣不安あり", "緊急確認必要"])
            with col2:
                key_hold = st.selectbox("鍵預かり", ["未確認", "なし", "あり", "検討中"])
                neighborhood = st.selectbox("近隣不安", ["未確認", "なし", "少しあり", "強くあり"])
                frequency = st.selectbox("管理頻度", ["未設定", "月1回", "月2回", "必要時", "一時確認のみ"])
                memo = st.text_area("メモ")

            submitted = st.form_submit_button("空き家カードを登録")

            if submitted:
                new_row = {
                    "property_id": make_id("prop"),
                    "case_id": case_id,
                    "client_id": client_id,
                    "登録日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "物件名": property_name,
                    "所在地": address,
                    "物件状態": property_status,
                    "空き家状態": vacant_status,
                    "鍵預かり": key_hold,
                    "近隣不安": neighborhood,
                    "管理頻度": frequency,
                    "メモ": memo,
                }
                data["properties"] = pd.concat([data["properties"], pd.DataFrame([new_row])], ignore_index=True)
                save_all(data)
                st.success("空き家カードを登録しました。")

        st.divider()
        st.dataframe(data["properties"][data["properties"]["case_id"] == case_id], use_container_width=True)


with tabs[7]:
    st.subheader("GoogleMap")
    st.caption("空き家カードの所在地からGoogleMapを開けます。APIキー不要の簡易版です。")

    if data["properties"].empty:
        st.info("空き家カードに所在地を登録すると、ここに表示されます。")
    else:
        for _, row in data["properties"].iterrows():
            address = row["所在地"]
            if address:
                map_url = f"https://www.google.com/maps/search/?api=1&query={quote_plus(address)}"
                with st.container(border=True):
                    st.write(f"物件名：{row['物件名']}")
                    st.write(f"所在地：{address}")
                    st.link_button("GoogleMapで開く", map_url)


with tabs[8]:
    st.subheader("猫情報カード")

    if data["cases"].empty:
        st.info("先に案件を登録してください。")
    else:
        case_labels = [get_case_label(row) for _, row in data["cases"].iterrows()]
        selected_case_label = st.selectbox("案件選択（猫情報カード）", case_labels, key="cat_select_case")
        case_id = selected_id_from_label(selected_case_label)
        case_row = data["cases"][data["cases"]["case_id"] == case_id].iloc[0]
        client_id = case_row["client_id"]

        with st.form("cat_form"):
            col1, col2 = st.columns(2)
            with col1:
                cat_name = st.text_input("猫の名前")
                cat_age = st.text_input("年齢")
                cat_count = st.text_input("頭数")
                current_life = st.selectbox("現在の暮らし", ["未確認", "本人と同居", "家族と同居", "一時預かり中", "今後検討", "その他"])
            with col2:
                cat_worry = st.text_area("気になること")
                place_candidate = st.text_area("預け先候補")
                cat_memo = st.text_area("メモ")

            submitted = st.form_submit_button("猫情報カードを登録")

            if submitted:
                new_row = {
                    "cat_id": make_id("cat"),
                    "case_id": case_id,
                    "client_id": client_id,
                    "登録日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "猫の名前": cat_name,
                    "年齢": cat_age,
                    "頭数": cat_count,
                    "現在の暮らし": current_life,
                    "気になること": cat_worry,
                    "預け先候補": place_candidate,
                    "メモ": cat_memo,
                }
                data["cats"] = pd.concat([data["cats"], pd.DataFrame([new_row])], ignore_index=True)
                save_all(data)
                st.success("猫情報カードを登録しました。")

        st.divider()
        st.dataframe(data["cats"][data["cats"]["case_id"] == case_id], use_container_width=True)


with tabs[9]:
    st.subheader("家族関係メモ")

    if data["cases"].empty:
        st.info("先に案件を登録してください。")
    else:
        case_labels = [get_case_label(row) for _, row in data["cases"].iterrows()]
        selected_case_label = st.selectbox("案件選択（家族関係メモ）", case_labels, key="family_select_case")
        case_id = selected_id_from_label(selected_case_label)
        case_row = data["cases"][data["cases"]["case_id"] == case_id].iloc[0]
        client_id = case_row["client_id"]

        with st.form("family_form"):
            col1, col2 = st.columns(2)
            with col1:
                person_name = st.text_input("関係者名")
                relation = st.text_input("続柄")
                contact_ok = st.selectbox("連絡可否", ["未確認", "連絡可", "連絡不可", "本人経由のみ"])
            with col2:
                temperature = st.selectbox("温度感", ["未確認", "協力的", "中立", "慎重", "反対気味", "不明"])
                relation_memo = st.text_area("関係メモ")

            submitted = st.form_submit_button("家族関係メモを登録")

            if submitted:
                new_row = {
                    "family_id": make_id("fam"),
                    "case_id": case_id,
                    "client_id": client_id,
                    "登録日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "関係者名": person_name,
                    "続柄": relation,
                    "連絡可否": contact_ok,
                    "温度感": temperature,
                    "関係メモ": relation_memo,
                }
                data["family"] = pd.concat([data["family"], pd.DataFrame([new_row])], ignore_index=True)
                save_all(data)
                st.success("家族関係メモを登録しました。")

        st.divider()
        st.dataframe(data["family"][data["family"]["case_id"] == case_id], use_container_width=True)


with tabs[10]:
    st.subheader("写真管理")
    st.caption("現地写真、猫写真、書類写真などを案件ごとに保存します。")

    if data["cases"].empty:
        st.info("先に案件を登録してください。")
    else:
        case_labels = [get_case_label(row) for _, row in data["cases"].iterrows()]
        selected_case_label = st.selectbox("案件選択（写真管理）", case_labels, key="photo_select_case")
        case_id = selected_id_from_label(selected_case_label)
        case_row = data["cases"][data["cases"]["case_id"] == case_id].iloc[0]
        client_id = case_row["client_id"]

        photo_type = st.selectbox("写真種別", ["外観", "室内", "郵便受け", "庭", "猫", "書類", "その他"])
        description = st.text_area("写真説明")
        uploaded_files = st.file_uploader(
            "写真をアップロード",
            type=["png", "jpg", "jpeg"],
            accept_multiple_files=True
        )

        if st.button("写真を保存"):
            if not uploaded_files:
                st.error("写真を選択してください。")
            else:
                rows = []
                for uploaded in uploaded_files:
                    suffix = Path(uploaded.name).suffix.lower()
                    photo_id = make_id("photo")
                    safe_name = f"{photo_id}{suffix}"
                    save_path = PHOTO_DIR / safe_name
                    save_path.write_bytes(uploaded.getbuffer())

                    rows.append({
                        "photo_id": photo_id,
                        "case_id": case_id,
                        "client_id": client_id,
                        "登録日時": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "写真種別": photo_type,
                        "ファイル名": uploaded.name,
                        "保存先": str(save_path),
                        "説明": description,
                    })

                data["photos"] = pd.concat([data["photos"], pd.DataFrame(rows)], ignore_index=True)
                save_all(data)
                st.success("写真を保存しました。")

        st.divider()
        photo_df = data["photos"][data["photos"]["case_id"] == case_id]
        st.dataframe(photo_df, use_container_width=True)

        for _, p in photo_df.iterrows():
            path = Path(p["保存先"])
            if path.exists():
                with st.container(border=True):
                    st.write(f"{p['写真種別']}：{p['ファイル名']}")
                    st.write(p["説明"])
                    st.image(str(path), width=350)


with tabs[11]:
    st.subheader("AI要約")
    st.caption("外部AIに送る前の安全な下書きと、コピー用プロンプトを作成します。")

    if data["cases"].empty:
        st.info("案件がありません。")
    else:
        case_labels = [get_case_label(row) for _, row in data["cases"].iterrows()]
        selected_case_label = st.selectbox("案件選択（AI要約）", case_labels, key="ai_select_case")
        case_id = selected_id_from_label(selected_case_label)

        memo = build_case_memo(data, case_id)
        summary = build_ai_summary(data, case_id)
        prompt = build_ai_prompt(memo)

        st.markdown("### 自動要約下書き")
        st.text_area("AI要約・下書き", summary, height=350)

        st.markdown("### AIに貼り付ける用プロンプト")
        st.text_area("コピー用プロンプト", prompt, height=500)

        st.warning("個人情報を外部AIへ入力する場合は、匿名化・伏せ字化してから使用してください。")


with tabs[12]:
    st.subheader("PDF出力")

    if data["cases"].empty:
        st.info("案件がありません。")
    else:
        case_labels = [get_case_label(row) for _, row in data["cases"].iterrows()]
        selected_case_label = st.selectbox("案件選択（PDF出力）", case_labels, key="pdf_select_case")
        case_id = selected_id_from_label(selected_case_label)

        memo = build_case_memo(data, case_id)
        st.text_area("PDF化する相談整理メモ", memo, height=500)

        pdf_bytes = make_pdf_bytes(memo)
        st.download_button(
            "PDFをダウンロード",
            data=pdf_bytes,
            file_name=f"nyantomo_memo_{case_id}.pdf",
            mime="application/pdf",
        )


with tabs[13]:
    data = render_search_update_delete(data)


with tabs[14]:
    st.subheader("データ管理")

    if DATA_FILE.exists():
        with open(DATA_FILE, "rb") as f:
            st.download_button(
                "Excelデータをダウンロード",
                data=f,
                file_name="nyantomo_consultation_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    st.markdown("### 全データ一覧")
    for label, cfg in TABLE_CONFIG.items():
        with st.expander(label):
            st.dataframe(data[cfg["key"]], use_container_width=True)

    st.markdown("### 注意")
    st.write(
        """
        Streamlit Cloudでは、無料環境の仕様により保存データが永続化されない場合があります。
        本格運用前には、SQLite化、クラウドDB化、または定期バックアップを行う前提で使ってください。
        """
    )
