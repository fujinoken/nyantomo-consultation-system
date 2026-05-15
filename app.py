
import streamlit as st
import pandas as pd
import sqlite3
from datetime import date, datetime
from pathlib import Path
import uuid
import io
import zipfile
import shutil
import json
from urllib.parse import quote_plus

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase import pdfmetrics


# =========================================================
# にゃんとも相談管理システム Ver2.0 SQLite版
# ---------------------------------------------------------
# 方針：
# ・client_id / case_id を正式な主キーとして管理
# ・相談者・案件・履歴・空き家・猫・家族・写真をSQLiteで保存
# ・Excel版からの移行にも対応
# ・診断しない／結論を急がせない／未確定と保留を分けて記録
# =========================================================

st.set_page_config(
    page_title="にゃんとも相談管理システム Ver2.0",
    page_icon="🐾",
    layout="wide"
)

DB_FILE = Path("nyantomo_consultation.db")
OLD_EXCEL_FILE = Path("nyantomo_consultation_data.xlsx")
PHOTO_DIR = Path("photos")
PHOTO_DIR.mkdir(exist_ok=True)


# -----------------------------
# 選択肢
# -----------------------------
STATUS_ORDER = [
    "未対応", "初回相談前", "初回相談済", "情報整理中", "保留中",
    "見守り中", "継続相談", "専門家紹介済", "終了"
]

CASE_TYPES = ["初回相談", "空き家管理", "猫と住まい", "相続前整理", "高齢期の住まい", "その他"]
AGE_OPTIONS = ["未選択", "40代", "50代", "60代", "70代", "80代以上"]
CONTACT_OPTIONS = ["未選択", "LINE", "メール", "電話", "対面", "その他"]
POSITION_OPTIONS = ["未選択", "本人", "家族", "親族", "空き家所有者", "支援者", "その他"]
CURRENT_STATE_OPTIONS = ["未選択", "まだ何も決まっていない", "少し考え始めている", "家族と話し始めた", "急かされている感じがある", "誰にも相談していない", "すでに困りごとが出ている"]
HOUSE_STATE_OPTIONS = ["未選択", "現在住んでいる", "空き家になっている", "近いうちに空き家になりそう", "相続後そのまま", "売却・賃貸を迷っている", "荷物整理が進んでいない"]
CAT_RELATION_OPTIONS = ["未選択", "猫と暮らしている", "家族の猫がいる", "猫を残して入院・施設入所が心配", "これから猫と暮らしたい", "保護猫に関心がある", "猫はいない"]
FAMILY_GAP_OPTIONS = ["未選択", "特にない", "少しある", "かなりある", "まだ話せていない"]
PRESSURE_OPTIONS = ["未選択", "ない", "少しある", "強くある", "自分でも焦っている"]
WORRY_OPTIONS = ["空き家管理", "相続", "売却", "賃貸", "猫の住まい", "高齢期の暮らし", "家族との意見の違い", "お金", "近所への不安", "何から考えればよいか分からない"]

PROPERTY_STATUS_OPTIONS = ["未確認", "居住中", "空き家", "一部使用", "売却検討", "賃貸検討", "その他"]
VACANT_STATUS_OPTIONS = ["未確認", "問題なし", "定期確認必要", "劣化あり", "近隣不安あり", "緊急確認必要"]
KEY_HOLD_OPTIONS = ["未確認", "なし", "あり", "検討中"]
NEIGHBORHOOD_OPTIONS = ["未確認", "なし", "少しあり", "強くあり"]
FREQUENCY_OPTIONS = ["未設定", "月1回", "月2回", "必要時", "一時確認のみ"]
CAT_LIFE_OPTIONS = ["未確認", "本人と同居", "家族と同居", "一時預かり中", "今後検討", "その他"]
CONTACT_OK_OPTIONS = ["未確認", "連絡可", "連絡不可", "本人経由のみ"]
TEMP_OPTIONS = ["未確認", "協力的", "中立", "慎重", "反対気味", "不明"]
PHOTO_TYPE_OPTIONS = ["外観", "室内", "郵便受け", "庭", "猫", "書類", "その他"]
REOPEN_OPTIONS = ["未確認", "低い", "あり得る", "高い", "定期確認予定"]


# -----------------------------
# 基本ユーティリティ
# -----------------------------
def make_id(prefix: str) -> str:
    return f"{prefix}_{uuid.uuid4().hex[:10]}"


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today_text() -> str:
    return date.today().strftime("%Y-%m-%d")


def get_conn():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def execute(sql, params=None):
    with get_conn() as conn:
        conn.execute(sql, params or {})
        conn.commit()


def fetch_df(sql, params=None):
    with get_conn() as conn:
        return pd.read_sql_query(sql, conn, params=params or {})


def fetch_one(sql, params=None):
    with get_conn() as conn:
        row = conn.execute(sql, params or {}).fetchone()
        return dict(row) if row else None


def list_to_text(values):
    if isinstance(values, list):
        return "、".join(values)
    return str(values or "")


def text_to_list(text):
    if not text:
        return []
    return [x for x in str(text).split("、") if x]


def date_or_blank(d):
    if not d:
        return ""
    if isinstance(d, date):
        return d.strftime("%Y-%m-%d")
    return str(d)


def parse_date_safe(value):
    try:
        if not value:
            return date.today()
        return pd.to_datetime(value).date()
    except Exception:
        return date.today()


def option_index(options, value):
    try:
        return options.index(value)
    except Exception:
        return 0


def selected_id_from_label(label):
    if not label or "｜" not in label:
        return ""
    return label.split("｜")[-1].strip()


# -----------------------------
# DB初期化・マイグレーション
# -----------------------------
def add_column_if_missing(conn, table, col, col_type="TEXT"):
    cols = [r["name"] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
    if col not in cols:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {col_type}")


def init_db():
    with get_conn() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS clients (
            client_id TEXT PRIMARY KEY,
            created_at TEXT,
            name TEXT NOT NULL,
            age_group TEXT,
            area TEXT,
            contact_method TEXT,
            position TEXT,
            note TEXT
        );

        CREATE TABLE IF NOT EXISTS cases (
            case_id TEXT PRIMARY KEY,
            client_id TEXT NOT NULL,
            created_at TEXT,
            consult_date TEXT,
            case_title TEXT,
            case_type TEXT,
            status TEXT,
            current_state TEXT,
            house_state TEXT,
            cat_relation TEXT,
            family_gap TEXT,
            pressure TEXT,
            worries TEXT,
            not_decide TEXT,
            first_check TEXT,
            free_memo TEXT,
            internal_memo TEXT,
            next_check TEXT,
            next_check_date TEXT,
            closed_date TEXT,
            close_reason TEXT,
            final_memo TEXT,
            reopen_possibility TEXT,
            updated_at TEXT,
            FOREIGN KEY(client_id) REFERENCES clients(client_id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS history (
            history_id TEXT PRIMARY KEY,
            case_id TEXT NOT NULL,
            client_id TEXT NOT NULL,
            created_at TEXT,
            record_date TEXT,
            record_type TEXT,
            before_status TEXT,
            after_status TEXT,
            record TEXT,
            next_action TEXT,
            internal_memo TEXT,
            FOREIGN KEY(case_id) REFERENCES cases(case_id) ON DELETE CASCADE,
            FOREIGN KEY(client_id) REFERENCES clients(client_id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS properties (
            property_id TEXT PRIMARY KEY,
            case_id TEXT NOT NULL,
            client_id TEXT NOT NULL,
            created_at TEXT,
            property_name TEXT,
            address TEXT,
            property_status TEXT,
            vacant_status TEXT,
            key_hold TEXT,
            neighborhood_anxiety TEXT,
            management_frequency TEXT,
            memo TEXT,
            FOREIGN KEY(case_id) REFERENCES cases(case_id) ON DELETE CASCADE,
            FOREIGN KEY(client_id) REFERENCES clients(client_id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS cats (
            cat_id TEXT PRIMARY KEY,
            case_id TEXT NOT NULL,
            client_id TEXT NOT NULL,
            created_at TEXT,
            cat_name TEXT,
            age TEXT,
            count TEXT,
            current_life TEXT,
            concerns TEXT,
            place_candidate TEXT,
            memo TEXT,
            FOREIGN KEY(case_id) REFERENCES cases(case_id) ON DELETE CASCADE,
            FOREIGN KEY(client_id) REFERENCES clients(client_id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS family (
            family_id TEXT PRIMARY KEY,
            case_id TEXT NOT NULL,
            client_id TEXT NOT NULL,
            created_at TEXT,
            person_name TEXT,
            relation TEXT,
            contact_ok TEXT,
            temperature TEXT,
            relation_memo TEXT,
            FOREIGN KEY(case_id) REFERENCES cases(case_id) ON DELETE CASCADE,
            FOREIGN KEY(client_id) REFERENCES clients(client_id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS photos (
            photo_id TEXT PRIMARY KEY,
            case_id TEXT NOT NULL,
            client_id TEXT NOT NULL,
            created_at TEXT,
            photo_type TEXT,
            original_filename TEXT,
            saved_path TEXT,
            description TEXT,
            FOREIGN KEY(case_id) REFERENCES cases(case_id) ON DELETE CASCADE,
            FOREIGN KEY(client_id) REFERENCES clients(client_id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS meta (
            key TEXT PRIMARY KEY,
            value TEXT
        );
        """)
        # 将来追加分に備えた軽いマイグレーション
        for col in ["next_check_date", "closed_date", "close_reason", "final_memo", "reopen_possibility", "updated_at"]:
            add_column_if_missing(conn, "cases", col)
        conn.execute("INSERT OR REPLACE INTO meta(key, value) VALUES('app_version', '2.0')")
        conn.commit()


# -----------------------------
# Excel版からSQLiteへ移行
# -----------------------------
EXCEL_MAP = {
    "clients": {
        "sheet": "clients",
        "cols": {
            "client_id": "client_id",
            "登録日時": "created_at",
            "お名前": "name",
            "年代": "age_group",
            "地域": "area",
            "連絡方法": "contact_method",
            "相談者の立場": "position",
            "備考": "note",
        }
    },
    "cases": {
        "sheet": "cases",
        "cols": {
            "case_id": "case_id",
            "client_id": "client_id",
            "登録日時": "created_at",
            "相談日": "consult_date",
            "案件名": "case_title",
            "案件種別": "case_type",
            "現在ステータス": "status",
            "今いちばん近い状態": "current_state",
            "住まいの状態": "house_state",
            "猫との関係": "cat_relation",
            "家族との温度差": "family_gap",
            "急がされている感じ": "pressure",
            "気になること": "worries",
            "今は決めたくないこと": "not_decide",
            "まず確認したいこと": "first_check",
            "自由メモ": "free_memo",
            "内部メモ": "internal_memo",
            "次回確認すること": "next_check",
            "次回確認日": "next_check_date",
            "終了日": "closed_date",
            "終了理由": "close_reason",
            "最終メモ": "final_memo",
            "再相談可能性": "reopen_possibility",
        }
    },
    "history": {
        "sheet": "history",
        "cols": {
            "history_id": "history_id",
            "case_id": "case_id",
            "client_id": "client_id",
            "記録日時": "created_at",
            "記録日": "record_date",
            "記録種別": "record_type",
            "状態変更前": "before_status",
            "状態変更後": "after_status",
            "相談記録": "record",
            "次回アクション": "next_action",
            "内部メモ": "internal_memo",
        }
    },
    "properties": {
        "sheet": "properties",
        "cols": {
            "property_id": "property_id",
            "case_id": "case_id",
            "client_id": "client_id",
            "登録日時": "created_at",
            "物件名": "property_name",
            "所在地": "address",
            "物件状態": "property_status",
            "空き家状態": "vacant_status",
            "鍵預かり": "key_hold",
            "近隣不安": "neighborhood_anxiety",
            "管理頻度": "management_frequency",
            "メモ": "memo",
        }
    },
    "cats": {
        "sheet": "cats",
        "cols": {
            "cat_id": "cat_id",
            "case_id": "case_id",
            "client_id": "client_id",
            "登録日時": "created_at",
            "猫の名前": "cat_name",
            "年齢": "age",
            "頭数": "count",
            "現在の暮らし": "current_life",
            "気になること": "concerns",
            "預け先候補": "place_candidate",
            "メモ": "memo",
        }
    },
    "family": {
        "sheet": "family",
        "cols": {
            "family_id": "family_id",
            "case_id": "case_id",
            "client_id": "client_id",
            "登録日時": "created_at",
            "関係者名": "person_name",
            "続柄": "relation",
            "連絡可否": "contact_ok",
            "温度感": "temperature",
            "関係メモ": "relation_memo",
        }
    },
    "photos": {
        "sheet": "photos",
        "cols": {
            "photo_id": "photo_id",
            "case_id": "case_id",
            "client_id": "client_id",
            "登録日時": "created_at",
            "写真種別": "photo_type",
            "ファイル名": "original_filename",
            "保存先": "saved_path",
            "説明": "description",
        }
    },
}


def table_count(table):
    row = fetch_one(f"SELECT COUNT(*) AS cnt FROM {table}")
    return int(row["cnt"]) if row else 0


def db_has_data():
    return sum(table_count(t) for t in ["clients", "cases", "history", "properties", "cats", "family", "photos"]) > 0


def import_excel_to_sqlite(excel_bytes=None):
    if excel_bytes:
        xls = pd.ExcelFile(io.BytesIO(excel_bytes))
    else:
        xls = pd.ExcelFile(OLD_EXCEL_FILE)

    imported = {}
    with get_conn() as conn:
        for table, cfg in EXCEL_MAP.items():
            if cfg["sheet"] not in xls.sheet_names:
                imported[table] = 0
                continue
            df = pd.read_excel(xls, sheet_name=cfg["sheet"], dtype=str).fillna("")
            if df.empty:
                imported[table] = 0
                continue

            new_rows = []
            for _, row in df.iterrows():
                mapped = {}
                for old_col, new_col in cfg["cols"].items():
                    mapped[new_col] = str(row.get(old_col, ""))

                # 必須ID補完
                if table == "clients" and not mapped.get("client_id"):
                    mapped["client_id"] = make_id("client")
                if table == "cases" and not mapped.get("case_id"):
                    mapped["case_id"] = make_id("case")
                if table == "history" and not mapped.get("history_id"):
                    mapped["history_id"] = make_id("hist")
                if table == "properties" and not mapped.get("property_id"):
                    mapped["property_id"] = make_id("prop")
                if table == "cats" and not mapped.get("cat_id"):
                    mapped["cat_id"] = make_id("cat")
                if table == "family" and not mapped.get("family_id"):
                    mapped["family_id"] = make_id("fam")
                if table == "photos" and not mapped.get("photo_id"):
                    mapped["photo_id"] = make_id("photo")

                if table == "cases" and not mapped.get("updated_at"):
                    mapped["updated_at"] = mapped.get("created_at", now_text())

                new_rows.append(mapped)

            if new_rows:
                cols = list(new_rows[0].keys())
                placeholders = ",".join([f":{c}" for c in cols])
                col_sql = ",".join(cols)
                conn.executemany(
                    f"INSERT OR IGNORE INTO {table} ({col_sql}) VALUES ({placeholders})",
                    new_rows
                )
                imported[table] = len(new_rows)
            else:
                imported[table] = 0
        conn.commit()
    return imported


# -----------------------------
# データ取得
# -----------------------------
def get_clients_df():
    return fetch_df("SELECT * FROM clients ORDER BY created_at DESC")


def get_cases_df(include_closed=True):
    where = "" if include_closed else "WHERE c.status != '終了'"
    return fetch_df(f"""
        SELECT
            c.case_id,
            c.client_id,
            cl.name AS 相談者,
            cl.area AS 地域,
            c.case_title AS 案件名,
            c.case_type AS 案件種別,
            c.status AS 現在ステータス,
            c.consult_date AS 相談日,
            c.next_check_date AS 次回確認日,
            c.next_check AS 次回確認すること,
            c.updated_at AS 最終更新,
            c.closed_date AS 終了日,
            c.close_reason AS 終了理由,
            c.reopen_possibility AS 再相談可能性,
            c.current_state AS 今いちばん近い状態,
            c.house_state AS 住まいの状態,
            c.cat_relation AS 猫との関係,
            c.family_gap AS 家族との温度差,
            c.pressure AS 急がされている感じ,
            c.worries AS 気になること
        FROM cases c
        JOIN clients cl ON c.client_id = cl.client_id
        {where}
        ORDER BY
            CASE WHEN c.status='終了' THEN 1 ELSE 0 END,
            COALESCE(c.next_check_date, '9999-12-31') ASC,
            c.updated_at DESC
    """)


def get_case_full(case_id):
    return fetch_one("""
        SELECT c.*, cl.name, cl.area, cl.age_group, cl.contact_method, cl.position, cl.note AS client_note
        FROM cases c
        JOIN clients cl ON c.client_id = cl.client_id
        WHERE c.case_id = :case_id
    """, {"case_id": case_id})


def get_case_options(include_closed=False):
    df = get_cases_df(include_closed=include_closed)
    labels = []
    for _, r in df.iterrows():
        labels.append(f"{r['相談者']}｜{r['案件名']}｜{r['現在ステータス']}｜{r.get('次回確認日','') or '日付未設定'}｜{r['case_id']}")
    return labels


def get_client_options():
    df = get_clients_df()
    return [f"{r['name']}｜{r['area']}｜{r['client_id']}" for _, r in df.iterrows()]


def get_related_df(table, case_id):
    return fetch_df(f"SELECT * FROM {table} WHERE case_id = :case_id ORDER BY created_at DESC", {"case_id": case_id})


def get_last_update_days(updated_at):
    if not updated_at:
        return None
    try:
        dt = pd.to_datetime(updated_at).date()
        return (date.today() - dt).days
    except Exception:
        return None


def get_input_shortage_count(case_row):
    keys = ["current_state", "house_state", "cat_relation", "family_gap", "pressure", "worries", "next_check", "next_check_date"]
    cnt = 0
    for k in keys:
        v = case_row.get(k, "")
        if not v or v == "未選択":
            cnt += 1
    return cnt


def get_today_tasks_df():
    today = today_text()
    return fetch_df("""
        SELECT
            c.case_id,
            cl.name AS 相談者,
            c.case_title AS 案件名,
            c.status AS 現在ステータス,
            c.next_check_date AS 次回確認日,
            c.next_check AS 次回確認すること,
            c.updated_at AS 最終更新
        FROM cases c
        JOIN clients cl ON c.client_id = cl.client_id
        WHERE c.status != '終了'
          AND c.next_check_date IS NOT NULL
          AND c.next_check_date != ''
          AND c.next_check_date <= :today
        ORDER BY c.next_check_date ASC
    """, {"today": today})


def get_soon_tasks_df(days=7):
    end = (date.today() + pd.Timedelta(days=days)).strftime("%Y-%m-%d")
    today = today_text()
    return fetch_df("""
        SELECT
            c.case_id,
            cl.name AS 相談者,
            c.case_title AS 案件名,
            c.status AS 現在ステータス,
            c.next_check_date AS 次回確認日,
            c.next_check AS 次回確認すること
        FROM cases c
        JOIN clients cl ON c.client_id = cl.client_id
        WHERE c.status != '終了'
          AND c.next_check_date IS NOT NULL
          AND c.next_check_date != ''
          AND c.next_check_date > :today
          AND c.next_check_date <= :end
        ORDER BY c.next_check_date ASC
    """, {"today": today, "end": end})


def get_stale_cases_df(days=60):
    df = get_cases_df(include_closed=False)
    rows = []
    for _, r in df.iterrows():
        d = get_last_update_days(r.get("最終更新", ""))
        if d is not None and d >= days:
            rr = r.copy()
            rr["未更新日数"] = d
            rows.append(rr)
    return pd.DataFrame(rows) if rows else pd.DataFrame()


# -----------------------------
# 登録・更新
# -----------------------------
def insert_client(name, age_group, area, contact_method, position, note):
    client_id = make_id("client")
    execute("""
        INSERT INTO clients(client_id, created_at, name, age_group, area, contact_method, position, note)
        VALUES(:client_id, :created_at, :name, :age_group, :area, :contact_method, :position, :note)
    """, {
        "client_id": client_id,
        "created_at": now_text(),
        "name": name,
        "age_group": age_group,
        "area": area,
        "contact_method": contact_method,
        "position": position,
        "note": note,
    })
    return client_id


def insert_case(values):
    case_id = make_id("case")
    now = now_text()
    values = dict(values)
    values.update({"case_id": case_id, "created_at": now, "updated_at": now})
    execute("""
        INSERT INTO cases(
            case_id, client_id, created_at, consult_date, case_title, case_type, status,
            current_state, house_state, cat_relation, family_gap, pressure, worries,
            not_decide, first_check, free_memo, internal_memo, next_check, next_check_date,
            updated_at
        )
        VALUES(
            :case_id, :client_id, :created_at, :consult_date, :case_title, :case_type, :status,
            :current_state, :house_state, :cat_relation, :family_gap, :pressure, :worries,
            :not_decide, :first_check, :free_memo, :internal_memo, :next_check, :next_check_date,
            :updated_at
        )
    """, values)

    insert_history({
        "case_id": case_id,
        "client_id": values["client_id"],
        "record_date": values["consult_date"],
        "record_type": "案件登録",
        "before_status": "",
        "after_status": values["status"],
        "record": values["free_memo"],
        "next_action": values["next_check"],
        "internal_memo": values["internal_memo"],
        "update_case": False,
    })
    return case_id


def insert_history(values):
    history_id = make_id("hist")
    params = {
        "history_id": history_id,
        "case_id": values["case_id"],
        "client_id": values["client_id"],
        "created_at": now_text(),
        "record_date": values.get("record_date", today_text()),
        "record_type": values.get("record_type", ""),
        "before_status": values.get("before_status", ""),
        "after_status": values.get("after_status", ""),
        "record": values.get("record", ""),
        "next_action": values.get("next_action", ""),
        "internal_memo": values.get("internal_memo", ""),
    }
    execute("""
        INSERT INTO history(history_id, case_id, client_id, created_at, record_date, record_type,
                            before_status, after_status, record, next_action, internal_memo)
        VALUES(:history_id, :case_id, :client_id, :created_at, :record_date, :record_type,
               :before_status, :after_status, :record, :next_action, :internal_memo)
    """, params)
    if values.get("update_case", True):
        execute("""
            UPDATE cases
            SET status=:status, next_check=:next_check, updated_at=:updated_at
            WHERE case_id=:case_id
        """, {
            "status": params["after_status"],
            "next_check": params["next_action"],
            "updated_at": now_text(),
            "case_id": params["case_id"],
        })
    return history_id


def update_case_basic(case_id, fields):
    fields = {k: v for k, v in fields.items()}
    fields["updated_at"] = now_text()
    fields["case_id"] = case_id
    set_sql = ", ".join([f"{k}=:{k}" for k in fields.keys() if k != "case_id"])
    execute(f"UPDATE cases SET {set_sql} WHERE case_id=:case_id", fields)


def close_case(case_id, closed_date, close_reason, final_memo, reopen_possibility):
    case = get_case_full(case_id)
    if not case:
        return
    execute("""
        UPDATE cases
        SET status='終了',
            closed_date=:closed_date,
            close_reason=:close_reason,
            final_memo=:final_memo,
            reopen_possibility=:reopen_possibility,
            updated_at=:updated_at
        WHERE case_id=:case_id
    """, {
        "case_id": case_id,
        "closed_date": closed_date,
        "close_reason": close_reason,
        "final_memo": final_memo,
        "reopen_possibility": reopen_possibility,
        "updated_at": now_text(),
    })
    insert_history({
        "case_id": case_id,
        "client_id": case["client_id"],
        "record_date": closed_date,
        "record_type": "終了確認",
        "before_status": case["status"],
        "after_status": "終了",
        "record": final_memo,
        "next_action": "",
        "internal_memo": f"終了理由：{close_reason}／再相談可能性：{reopen_possibility}",
        "update_case": False,
    })


def client_has_open_case(client_id):
    df = fetch_df("""
        SELECT c.case_id, c.case_title, c.status, c.consult_date, cl.name, cl.area
        FROM cases c
        JOIN clients cl ON c.client_id = cl.client_id
        WHERE c.client_id=:client_id AND c.status!='終了'
        ORDER BY c.created_at DESC
    """, {"client_id": client_id})
    return not df.empty, df


# -----------------------------
# PDF・バックアップ
# -----------------------------
def build_case_memo(case_id):
    c = get_case_full(case_id)
    if not c:
        return ""

    h = get_related_df("history", case_id)
    p = get_related_df("properties", case_id)
    cats = get_related_df("cats", case_id)
    fam = get_related_df("family", case_id)
    photos = get_related_df("photos", case_id)

    def rows_text(df, formatter):
        if df.empty:
            return "未登録"
        return "\n".join([formatter(row) for _, row in df.iterrows()]) or "未登録"

    history_text = rows_text(
        h,
        lambda r: f"- {r['record_date']}｜{r['record_type']}｜{r['before_status']} → {r['after_status']}｜{r['record']}"
    )
    prop_text = rows_text(
        p,
        lambda r: f"- {r['property_name']}｜{r['address']}｜{r['property_status']}｜{r['vacant_status']}｜管理頻度：{r['management_frequency']}"
    )
    cat_text = rows_text(
        cats,
        lambda r: f"- {r['cat_name']}｜年齢：{r['age']}｜頭数：{r['count']}｜{r['current_life']}｜気になること：{r['concerns']}"
    )
    fam_text = rows_text(
        fam,
        lambda r: f"- {r['person_name']}｜{r['relation']}｜{r['temperature']}｜{r['relation_memo']}"
    )
    photo_text = rows_text(
        photos,
        lambda r: f"- {r['photo_type']}｜{r['original_filename']}｜{r['description']}"
    )

    return f"""【にゃんとも相談整理メモ】

作成日：{today_text()}

■ 相談者
お名前：{c.get('name', '')}
地域：{c.get('area', '')}
年代：{c.get('age_group', '')}
連絡方法：{c.get('contact_method', '')}
相談者の立場：{c.get('position', '')}

■ 案件
案件名：{c.get('case_title', '')}
案件種別：{c.get('case_type', '')}
現在ステータス：{c.get('status', '')}
相談日：{c.get('consult_date', '')}
次回確認日：{c.get('next_check_date', '')}

■ 今いちばん近い状態
{c.get('current_state', '')}

■ 住まいの状態
{c.get('house_state', '')}

■ 猫との関係
{c.get('cat_relation', '')}

■ 気になること
{c.get('worries', '')}

■ 家族との温度差
{c.get('family_gap', '')}

■ 急がされている感じ
{c.get('pressure', '')}

■ 今は決めたくないこと
{c.get('not_decide', '')}

■ まず確認したいこと
{c.get('first_check', '')}

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
{c.get('internal_memo', '')}

■ 次回確認すること
{c.get('next_check', '')}

■ 終了情報
終了日：{c.get('closed_date', '')}
終了理由：{c.get('close_reason', '')}
最終メモ：{c.get('final_memo', '')}
再相談可能性：{c.get('reopen_possibility', '')}

※このメモは、判断を急がせず、状況を整理するための内部記録です。
※法的判断・医療判断・不動産判断を断定するものではありません。
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


def make_backup_zip_bytes():
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        if DB_FILE.exists():
            z.write(DB_FILE, DB_FILE.name)
        if PHOTO_DIR.exists():
            for path in PHOTO_DIR.rglob("*"):
                if path.is_file():
                    z.write(path, path.as_posix())
        meta = {
            "app": "nyantomo-consultation-system",
            "version": "2.0",
            "created_at": now_text(),
            "files": [DB_FILE.name, "photos/"]
        }
        z.writestr("backup_meta.json", json.dumps(meta, ensure_ascii=False, indent=2))
    buffer.seek(0)
    return buffer.getvalue()


def restore_backup_zip(uploaded_file):
    temp_dir = Path("_restore_temp")
    if temp_dir.exists():
        shutil.rmtree(temp_dir)
    temp_dir.mkdir(exist_ok=True)

    with zipfile.ZipFile(uploaded_file, "r") as z:
        z.extractall(temp_dir)

    restored = []
    db_path = temp_dir / DB_FILE.name
    if db_path.exists():
        shutil.copy2(db_path, DB_FILE)
        restored.append(DB_FILE.name)

    photos_path = temp_dir / "photos"
    if photos_path.exists():
        PHOTO_DIR.mkdir(exist_ok=True)
        for p in photos_path.rglob("*"):
            if p.is_file():
                dest = PHOTO_DIR / p.relative_to(photos_path)
                dest.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(p, dest)
        restored.append("photos")

    shutil.rmtree(temp_dir, ignore_errors=True)
    return restored


def export_all_to_excel_bytes():
    buffer = io.BytesIO()
    sheets = {
        "clients": fetch_df("SELECT * FROM clients"),
        "cases": fetch_df("SELECT * FROM cases"),
        "history": fetch_df("SELECT * FROM history"),
        "properties": fetch_df("SELECT * FROM properties"),
        "cats": fetch_df("SELECT * FROM cats"),
        "family": fetch_df("SELECT * FROM family"),
        "photos": fetch_df("SELECT * FROM photos"),
    }
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
    buffer.seek(0)
    return buffer.getvalue()


# -----------------------------
# UI補助
# -----------------------------
def render_top_nav():
    st.markdown("""
    <style>
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        flex-wrap: wrap;
    }
    .stTabs [data-baseweb="tab"] {
        height: 38px;
        white-space: nowrap;
        border-radius: 12px 12px 0 0;
    }
    .small-caption { color:#666; font-size:0.9rem; }
    </style>
    """, unsafe_allow_html=True)


def case_priority(row):
    if row.get("現在ステータス") == "終了":
        return "終了"
    if row.get("次回確認日"):
        try:
            d = pd.to_datetime(row.get("次回確認日")).date()
            if d < date.today():
                return "期限超過"
            if d == date.today():
                return "今日確認"
            if d <= date.today() + pd.Timedelta(days=7):
                return "7日以内"
        except Exception:
            pass
    shortage = 0
    for k in ["今いちばん近い状態", "住まいの状態", "猫との関係", "家族との温度差", "急がされている感じ", "気になること"]:
        v = row.get(k, "")
        if not v or v == "未選択":
            shortage += 1
    if shortage >= 3:
        return "入力不足"
    return "確認"


def decorate_cases_df(df):
    if df.empty:
        return df
    out = df.copy()
    out.insert(0, "優先", out.apply(case_priority, axis=1))
    return out


def select_case_widget(key, include_closed=False):
    options = get_case_options(include_closed=include_closed)
    if not options:
        st.info("案件がありません。")
        return ""
    label = st.selectbox("案件を選択", options, key=key)
    return selected_id_from_label(label)


# -----------------------------
# 各画面
# -----------------------------
def page_case_home():
    st.subheader("🏠 案件ホーム Ver2.0")
    st.caption("SQLite版：今日見る案件・期限・入力不足・未更新を確認します。")

    cases = get_cases_df(include_closed=True)
    open_cases = cases[cases["現在ステータス"] != "終了"] if not cases.empty else pd.DataFrame()
    today_df = get_today_tasks_df()
    soon_df = get_soon_tasks_df(7)
    stale_df = get_stale_cases_df(60)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("全案件", len(cases))
    c2.metric("進行中", len(open_cases))
    c3.metric("今日・期限超過", len(today_df))
    c4.metric("7日以内", len(soon_df))
    c5.metric("60日以上未更新", len(stale_df))

    st.markdown("### 今日やること")
    if today_df.empty:
        st.success("今日までに確認すべき案件はありません。")
    else:
        st.dataframe(today_df, use_container_width=True)

    with st.expander("7日以内に確認する案件", expanded=False):
        st.dataframe(soon_df, use_container_width=True)

    with st.expander("60日以上未更新の案件", expanded=False):
        st.dataframe(stale_df, use_container_width=True)

    st.divider()
    st.markdown("### 案件一覧")
    only_open = st.checkbox("終了以外を中心に見る", value=True, key="home_only_open")
    df = get_cases_df(include_closed=not only_open)
    keyword = st.text_input("案件ホーム検索", placeholder="案件名・相談者・次回確認など", key="home_keyword")
    if keyword:
        mask = df.astype(str).apply(lambda row: row.str.contains(keyword, case=False, na=False).any(), axis=1)
        df = df[mask]
    status_filter = st.multiselect("ステータス絞り込み", STATUS_ORDER, key="home_status_filter")
    if status_filter:
        df = df[df["現在ステータス"].isin(status_filter)]
    st.dataframe(decorate_cases_df(df), use_container_width=True)

    st.info("Ver2.0では、Excel保存ではなくSQLite DBで主キー・関連データを管理しています。")


def page_client_register():
    st.subheader("🧑 相談者登録")
    with st.form("client_form_v2"):
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("お名前")
            age = st.selectbox("年代", AGE_OPTIONS)
            area = st.text_input("地域")
        with col2:
            contact = st.selectbox("連絡方法", CONTACT_OPTIONS)
            position = st.selectbox("相談者の立場", POSITION_OPTIONS)
            note = st.text_area("備考")
        submitted = st.form_submit_button("相談者を登録")
        if submitted:
            if not name.strip():
                st.error("お名前を入力してください。")
            else:
                insert_client(name.strip(), age, area, contact, position, note)
                st.success("相談者を登録しました。")
                st.rerun()

    st.divider()
    st.dataframe(get_clients_df(), use_container_width=True)


def page_case_register():
    st.subheader("📝 案件登録")
    clients = get_clients_df()
    if clients.empty:
        st.info("先に相談者を登録してください。")
        return

    client_label = st.selectbox("相談者を選択", get_client_options(), key="case_reg_client_select")
    client_id = selected_id_from_label(client_label)

    has_open, open_df = client_has_open_case(client_id)
    if has_open:
        st.warning("この相談者には、すでに終了していない案件があります。二重登録防止のため、新規案件登録はできません。既存案件の「案件ダッシュボード」または「相談履歴」から追記してください。")
        st.dataframe(open_df, use_container_width=True)
        return

    with st.form("case_form_v2"):
        col1, col2 = st.columns(2)
        with col1:
            consult_date = st.date_input("相談日", value=date.today())
            case_title = st.text_input("案件名", value="住まいと猫の相談")
            case_type = st.selectbox("案件種別", CASE_TYPES)
            status = st.selectbox("現在ステータス", STATUS_ORDER, index=1)
            current_state = st.selectbox("今いちばん近い状態", CURRENT_STATE_OPTIONS)
            next_check_date = st.date_input("次回確認日", value=date.today())
        with col2:
            house_state = st.selectbox("住まいの状態", HOUSE_STATE_OPTIONS)
            cat_relation = st.selectbox("猫との関係", CAT_RELATION_OPTIONS)
            family_gap = st.selectbox("家族との温度差", FAMILY_GAP_OPTIONS)
            pressure = st.selectbox("急がされている感じ", PRESSURE_OPTIONS)

        worries = st.multiselect("気になること", WORRY_OPTIONS)
        not_decide = st.text_area("今は決めたくないこと")
        first_check = st.text_area("まず確認したいこと")
        free_memo = st.text_area("自由メモ")
        internal_memo = st.text_area("内部メモ")
        next_check = st.text_area("次回確認すること")

        submitted = st.form_submit_button("案件を登録")
        if submitted:
            case_id = insert_case({
                "client_id": client_id,
                "consult_date": date_or_blank(consult_date),
                "case_title": case_title,
                "case_type": case_type,
                "status": status,
                "current_state": current_state,
                "house_state": house_state,
                "cat_relation": cat_relation,
                "family_gap": family_gap,
                "pressure": pressure,
                "worries": list_to_text(worries),
                "not_decide": not_decide,
                "first_check": first_check,
                "free_memo": free_memo,
                "internal_memo": internal_memo,
                "next_check": next_check,
                "next_check_date": date_or_blank(next_check_date),
            })
            st.success(f"案件を登録しました：{case_id}")
            st.rerun()

    st.divider()
    st.dataframe(get_cases_df(include_closed=False), use_container_width=True)


def page_case_dashboard():
    st.subheader("🗂 案件ダッシュボード")
    case_id = select_case_widget("dash_case_select", include_closed=True)
    if not case_id:
        return

    c = get_case_full(case_id)
    if not c:
        st.error("案件が見つかりません。")
        return

    progress = 0
    if c.get("status") in STATUS_ORDER:
        progress = int((STATUS_ORDER.index(c["status"]) + 1) / len(STATUS_ORDER) * 100)

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("現在ステータス", c.get("status", ""))
    col2.metric("相談者", c.get("name", ""))
    col3.metric("案件種別", c.get("case_type", ""))
    col4.metric("終了までの目安", f"{progress}%")
    st.progress(progress)

    with st.container(border=True):
        st.markdown("### 案件の現在地")
        a, b = st.columns(2)
        with a:
            st.write(f"**案件名：** {c.get('case_title','')}")
            st.write(f"**相談日：** {c.get('consult_date','')}")
            st.write(f"**次回確認日：** {c.get('next_check_date','')}")
            st.write(f"**今いちばん近い状態：** {c.get('current_state','')}")
            st.write(f"**住まいの状態：** {c.get('house_state','')}")
        with b:
            st.write(f"**猫との関係：** {c.get('cat_relation','')}")
            st.write(f"**家族との温度差：** {c.get('family_gap','')}")
            st.write(f"**急がされている感じ：** {c.get('pressure','')}")
            st.write(f"**気になること：** {c.get('worries','')}")
            st.write(f"**次回確認：** {c.get('next_check','')}")

    st.markdown("### 状態を進める")
    with st.form(f"dash_history_form_{case_id}"):
        a, b = st.columns(2)
        with a:
            record_date = st.date_input("記録日", value=date.today(), key=f"dash_rec_date_{case_id}")
            record_type = st.selectbox("記録種別", ["状態変更", "相談", "電話", "LINE", "メール", "面談", "現地確認", "終了確認", "その他"], key=f"dash_rec_type_{case_id}")
            new_status = st.selectbox("新しいステータス", STATUS_ORDER, index=option_index(STATUS_ORDER, c.get("status", "")), key=f"dash_status_{case_id}")
        with b:
            next_date = st.date_input("次回確認日", value=parse_date_safe(c.get("next_check_date")), key=f"dash_next_date_{case_id}")
            next_action = st.text_area("次回アクション", value=c.get("next_check", ""), key=f"dash_next_action_{case_id}")
        record = st.text_area("相談記録・判断保留の理由・確認した事実", key=f"dash_record_{case_id}")
        internal = st.text_area("内部メモ", key=f"dash_internal_{case_id}")
        submitted = st.form_submit_button("履歴を追加してステータス更新")
        if submitted:
            insert_history({
                "case_id": case_id,
                "client_id": c["client_id"],
                "record_date": date_or_blank(record_date),
                "record_type": record_type,
                "before_status": c.get("status", ""),
                "after_status": new_status,
                "record": record,
                "next_action": next_action,
                "internal_memo": internal,
            })
            update_case_basic(case_id, {"next_check_date": date_or_blank(next_date)})
            st.success("履歴を追加し、案件ステータスを更新しました。")
            st.rerun()

    st.divider()
    st.markdown("### 終了処理")
    if c.get("status") == "終了":
        st.success(f"この案件は終了しています。終了日：{c.get('closed_date','')}")
        st.write(f"終了理由：{c.get('close_reason','')}")
        st.write(f"再相談可能性：{c.get('reopen_possibility','')}")
    else:
        with st.expander("この案件を終了する", expanded=False):
            with st.form(f"close_case_form_{case_id}"):
                closed_date = st.date_input("終了日", value=date.today(), key=f"close_date_{case_id}")
                close_reason = st.text_input("終了理由", placeholder="例：相談整理完了、専門家紹介済、本人希望により終了")
                final_memo = st.text_area("最終メモ")
                reopen = st.selectbox("再相談可能性", REOPEN_OPTIONS)
                ok = st.checkbox("終了処理を行います", key=f"close_ok_{case_id}")
                submitted = st.form_submit_button("終了する", disabled=not ok)
                if submitted:
                    close_case(case_id, date_or_blank(closed_date), close_reason, final_memo, reopen)
                    st.success("案件を終了しました。")
                    st.rerun()

    st.divider()
    st.markdown("### この案件に紐づくデータ")
    rel_tabs = st.tabs(["相談者", "相談履歴", "空き家", "猫", "家族", "写真", "AI/PDF用メモ"])

    with rel_tabs[0]:
        st.dataframe(fetch_df("SELECT * FROM clients WHERE client_id=:client_id", {"client_id": c["client_id"]}), use_container_width=True)

    with rel_tabs[1]:
        st.dataframe(get_related_df("history", case_id), use_container_width=True)

    with rel_tabs[2]:
        p_df = get_related_df("properties", case_id)
        st.dataframe(p_df, use_container_width=True)
        for _, r in p_df.iterrows():
            address = r.get("address", "")
            if address:
                st.link_button(f"GoogleMapで開く：{r.get('property_name','物件')}", f"https://www.google.com/maps/search/?api=1&query={quote_plus(address)}")

    with rel_tabs[3]:
        st.dataframe(get_related_df("cats", case_id), use_container_width=True)

    with rel_tabs[4]:
        st.dataframe(get_related_df("family", case_id), use_container_width=True)

    with rel_tabs[5]:
        photo_df = get_related_df("photos", case_id)
        st.dataframe(photo_df, use_container_width=True)
        for _, p in photo_df.iterrows():
            path = Path(str(p.get("saved_path", "")))
            if path.exists():
                with st.container(border=True):
                    st.write(f"{p.get('photo_type','')}：{p.get('original_filename','')}")
                    st.write(p.get("description", ""))
                    st.image(str(path), width=350)

    with rel_tabs[6]:
        memo = build_case_memo(case_id)
        st.text_area("案件統合メモ", memo, height=500)
        st.download_button("この案件のPDFをダウンロード", make_pdf_bytes(memo), file_name=f"nyantomo_case_{case_id}.pdf", mime="application/pdf")


def page_history():
    st.subheader("📚 相談履歴")
    case_id = select_case_widget("history_case_select", include_closed=True)
    if not case_id:
        return
    c = get_case_full(case_id)

    with st.form(f"history_form_{case_id}"):
        a, b = st.columns(2)
        with a:
            record_date = st.date_input("記録日", value=date.today(), key=f"hist_date_{case_id}")
            record_type = st.selectbox("記録種別", ["相談", "電話", "LINE", "メール", "面談", "現地確認", "状態変更", "その他"], key=f"hist_type_{case_id}")
        with b:
            new_status = st.selectbox("状態変更後", STATUS_ORDER, index=option_index(STATUS_ORDER, c.get("status", "")), key=f"hist_status_{case_id}")
            next_date = st.date_input("次回確認日", value=parse_date_safe(c.get("next_check_date")), key=f"hist_next_date_{case_id}")
        record = st.text_area("相談記録", key=f"hist_record_{case_id}")
        next_action = st.text_area("次回アクション", value=c.get("next_check", ""), key=f"hist_next_{case_id}")
        internal = st.text_area("内部メモ", key=f"hist_internal_{case_id}")
        submitted = st.form_submit_button("履歴を追加")
        if submitted:
            insert_history({
                "case_id": case_id,
                "client_id": c["client_id"],
                "record_date": date_or_blank(record_date),
                "record_type": record_type,
                "before_status": c.get("status", ""),
                "after_status": new_status,
                "record": record,
                "next_action": next_action,
                "internal_memo": internal,
            })
            update_case_basic(case_id, {"next_check_date": date_or_blank(next_date)})
            st.success("相談履歴を追加しました。")
            st.rerun()

    st.dataframe(get_related_df("history", case_id), use_container_width=True)


def page_hold_list():
    st.subheader("⏸ 保留案件一覧")
    df = get_cases_df(include_closed=False)
    hold_df = df[df["現在ステータス"].isin(["保留中", "情報整理中", "見守り中"])] if not df.empty else pd.DataFrame()
    if hold_df.empty:
        st.success("現在、保留・見守り中の案件はありません。")
    else:
        st.dataframe(decorate_cases_df(hold_df), use_container_width=True)


def page_property():
    st.subheader("🏠 空き家カード")
    case_id = select_case_widget("property_case_select", include_closed=False)
    if not case_id:
        return
    c = get_case_full(case_id)

    with st.form(f"property_form_{case_id}"):
        a, b = st.columns(2)
        with a:
            property_name = st.text_input("物件名")
            address = st.text_input("所在地")
            property_status = st.selectbox("物件状態", PROPERTY_STATUS_OPTIONS)
            vacant_status = st.selectbox("空き家状態", VACANT_STATUS_OPTIONS)
        with b:
            key_hold = st.selectbox("鍵預かり", KEY_HOLD_OPTIONS)
            neighborhood = st.selectbox("近隣不安", NEIGHBORHOOD_OPTIONS)
            frequency = st.selectbox("管理頻度", FREQUENCY_OPTIONS)
            memo = st.text_area("メモ")
        submitted = st.form_submit_button("空き家カードを登録")
        if submitted:
            execute("""
                INSERT INTO properties(property_id, case_id, client_id, created_at, property_name, address,
                                       property_status, vacant_status, key_hold, neighborhood_anxiety,
                                       management_frequency, memo)
                VALUES(:property_id, :case_id, :client_id, :created_at, :property_name, :address,
                       :property_status, :vacant_status, :key_hold, :neighborhood_anxiety,
                       :management_frequency, :memo)
            """, {
                "property_id": make_id("prop"),
                "case_id": case_id,
                "client_id": c["client_id"],
                "created_at": now_text(),
                "property_name": property_name,
                "address": address,
                "property_status": property_status,
                "vacant_status": vacant_status,
                "key_hold": key_hold,
                "neighborhood_anxiety": neighborhood,
                "management_frequency": frequency,
                "memo": memo,
            })
            update_case_basic(case_id, {})
            st.success("空き家カードを登録しました。")
            st.rerun()

    st.dataframe(get_related_df("properties", case_id), use_container_width=True)


def page_map():
    st.subheader("🗺 GoogleMap")
    df = fetch_df("""
        SELECT p.*, c.case_title, cl.name
        FROM properties p
        JOIN cases c ON p.case_id = c.case_id
        JOIN clients cl ON p.client_id = cl.client_id
        ORDER BY p.created_at DESC
    """)
    if df.empty:
        st.info("空き家カードに所在地を登録すると、ここに表示されます。")
        return

    for _, row in df.iterrows():
        address = row.get("address", "")
        if address:
            with st.container(border=True):
                st.write(f"相談者：{row.get('name','')} ／ 案件：{row.get('case_title','')}")
                st.write(f"物件名：{row.get('property_name','')}")
                st.write(f"所在地：{address}")
                st.link_button("GoogleMapで開く", f"https://www.google.com/maps/search/?api=1&query={quote_plus(address)}")


def page_cat():
    st.subheader("🐈 猫情報カード")
    case_id = select_case_widget("cat_case_select", include_closed=False)
    if not case_id:
        return
    c = get_case_full(case_id)

    with st.form(f"cat_form_{case_id}"):
        a, b = st.columns(2)
        with a:
            cat_name = st.text_input("猫の名前")
            cat_age = st.text_input("年齢")
            cat_count = st.text_input("頭数")
            current_life = st.selectbox("現在の暮らし", CAT_LIFE_OPTIONS)
        with b:
            concerns = st.text_area("気になること")
            place_candidate = st.text_area("預け先候補")
            memo = st.text_area("メモ")
        submitted = st.form_submit_button("猫情報カードを登録")
        if submitted:
            execute("""
                INSERT INTO cats(cat_id, case_id, client_id, created_at, cat_name, age, count,
                                 current_life, concerns, place_candidate, memo)
                VALUES(:cat_id, :case_id, :client_id, :created_at, :cat_name, :age, :count,
                       :current_life, :concerns, :place_candidate, :memo)
            """, {
                "cat_id": make_id("cat"),
                "case_id": case_id,
                "client_id": c["client_id"],
                "created_at": now_text(),
                "cat_name": cat_name,
                "age": cat_age,
                "count": cat_count,
                "current_life": current_life,
                "concerns": concerns,
                "place_candidate": place_candidate,
                "memo": memo,
            })
            update_case_basic(case_id, {})
            st.success("猫情報カードを登録しました。")
            st.rerun()

    st.dataframe(get_related_df("cats", case_id), use_container_width=True)


def page_family():
    st.subheader("👪 家族関係メモ")
    case_id = select_case_widget("family_case_select", include_closed=False)
    if not case_id:
        return
    c = get_case_full(case_id)

    with st.form(f"family_form_{case_id}"):
        a, b = st.columns(2)
        with a:
            person_name = st.text_input("関係者名")
            relation = st.text_input("続柄")
            contact_ok = st.selectbox("連絡可否", CONTACT_OK_OPTIONS)
        with b:
            temperature = st.selectbox("温度感", TEMP_OPTIONS)
            relation_memo = st.text_area("関係メモ")
        submitted = st.form_submit_button("家族関係メモを登録")
        if submitted:
            execute("""
                INSERT INTO family(family_id, case_id, client_id, created_at, person_name, relation,
                                   contact_ok, temperature, relation_memo)
                VALUES(:family_id, :case_id, :client_id, :created_at, :person_name, :relation,
                       :contact_ok, :temperature, :relation_memo)
            """, {
                "family_id": make_id("fam"),
                "case_id": case_id,
                "client_id": c["client_id"],
                "created_at": now_text(),
                "person_name": person_name,
                "relation": relation,
                "contact_ok": contact_ok,
                "temperature": temperature,
                "relation_memo": relation_memo,
            })
            update_case_basic(case_id, {})
            st.success("家族関係メモを登録しました。")
            st.rerun()

    st.dataframe(get_related_df("family", case_id), use_container_width=True)


def page_photos():
    st.subheader("📷 写真管理")
    case_id = select_case_widget("photo_case_select", include_closed=False)
    if not case_id:
        return
    c = get_case_full(case_id)

    photo_type = st.selectbox("写真種別", PHOTO_TYPE_OPTIONS)
    description = st.text_area("写真説明")
    uploaded_files = st.file_uploader("写真をアップロード", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

    if st.button("写真を保存"):
        if not uploaded_files:
            st.error("写真を選択してください。")
        else:
            for uploaded in uploaded_files:
                suffix = Path(uploaded.name).suffix.lower()
                photo_id = make_id("photo")
                safe_name = f"{photo_id}{suffix}"
                save_path = PHOTO_DIR / safe_name
                save_path.write_bytes(uploaded.getbuffer())
                execute("""
                    INSERT INTO photos(photo_id, case_id, client_id, created_at, photo_type,
                                       original_filename, saved_path, description)
                    VALUES(:photo_id, :case_id, :client_id, :created_at, :photo_type,
                           :original_filename, :saved_path, :description)
                """, {
                    "photo_id": photo_id,
                    "case_id": case_id,
                    "client_id": c["client_id"],
                    "created_at": now_text(),
                    "photo_type": photo_type,
                    "original_filename": uploaded.name,
                    "saved_path": str(save_path),
                    "description": description,
                })
            update_case_basic(case_id, {})
            st.success("写真を保存しました。")
            st.rerun()

    photo_df = get_related_df("photos", case_id)
    st.dataframe(photo_df, use_container_width=True)
    for _, p in photo_df.iterrows():
        path = Path(str(p.get("saved_path", "")))
        if path.exists():
            with st.container(border=True):
                st.write(f"{p.get('photo_type','')}：{p.get('original_filename','')}")
                st.write(p.get("description", ""))
                st.image(str(path), width=350)


def page_ai_pdf():
    st.subheader("🤖 AI要約・PDF出力")
    case_id = select_case_widget("ai_pdf_case_select", include_closed=True)
    if not case_id:
        return
    memo = build_case_memo(case_id)

    prompt = f"""あなたは、にゃんとも相談管理システムの記録整理係です。
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
    st.text_area("案件統合メモ", memo, height=420)
    st.text_area("AIに貼り付ける用プロンプト", prompt, height=420)
    st.warning("個人情報を外部AIへ入力する場合は、匿名化・伏せ字化してから使用してください。")

    st.download_button("PDFをダウンロード", make_pdf_bytes(memo), file_name=f"nyantomo_memo_{case_id}.pdf", mime="application/pdf")


def page_search_update_delete():
    st.subheader("🔎 検索・更新・削除")
    st.caption("Ver2.0ではSQLiteの各テーブルを検索できます。更新は案件を中心に最小限にしています。")

    table_map = {
        "相談者": "clients",
        "案件": "cases",
        "相談履歴": "history",
        "空き家": "properties",
        "猫": "cats",
        "家族": "family",
        "写真": "photos",
    }
    table_label = st.selectbox("対象データ", list(table_map.keys()))
    table = table_map[table_label]
    df = fetch_df(f"SELECT * FROM {table}")
    keyword = st.text_input("検索キーワード", placeholder="名前・住所・猫の名前・メモなど")
    if keyword and not df.empty:
        df = df[df.astype(str).apply(lambda row: row.str.contains(keyword, case=False, na=False).any(), axis=1)]
    st.write(f"検索結果：{len(df)}件")
    st.dataframe(df, use_container_width=True)

    if table == "cases" and not df.empty:
        st.markdown("### 案件の基本情報を更新")
        labels = []
        for _, r in df.iterrows():
            labels.append(f"{r.get('case_title','')}｜{r.get('status','')}｜{r.get('case_id','')}")
        selected = st.selectbox("更新する案件", labels, key="edit_case_select")
        case_id = selected_id_from_label(selected)
        c = get_case_full(case_id)
        with st.form(f"edit_case_form_{case_id}"):
            a, b = st.columns(2)
            with a:
                case_title = st.text_input("案件名", value=c.get("case_title", ""))
                case_type = st.selectbox("案件種別", CASE_TYPES, index=option_index(CASE_TYPES, c.get("case_type", "")))
                status = st.selectbox("現在ステータス", STATUS_ORDER, index=option_index(STATUS_ORDER, c.get("status", "")))
                consult_date = st.date_input("相談日", value=parse_date_safe(c.get("consult_date")))
                next_check_date = st.date_input("次回確認日", value=parse_date_safe(c.get("next_check_date")))
            with b:
                current_state = st.selectbox("今いちばん近い状態", CURRENT_STATE_OPTIONS, index=option_index(CURRENT_STATE_OPTIONS, c.get("current_state", "")))
                house_state = st.selectbox("住まいの状態", HOUSE_STATE_OPTIONS, index=option_index(HOUSE_STATE_OPTIONS, c.get("house_state", "")))
                cat_relation = st.selectbox("猫との関係", CAT_RELATION_OPTIONS, index=option_index(CAT_RELATION_OPTIONS, c.get("cat_relation", "")))
                family_gap = st.selectbox("家族との温度差", FAMILY_GAP_OPTIONS, index=option_index(FAMILY_GAP_OPTIONS, c.get("family_gap", "")))
                pressure = st.selectbox("急がされている感じ", PRESSURE_OPTIONS, index=option_index(PRESSURE_OPTIONS, c.get("pressure", "")))
            worries = st.multiselect("気になること", WORRY_OPTIONS, default=[x for x in text_to_list(c.get("worries", "")) if x in WORRY_OPTIONS])
            not_decide = st.text_area("今は決めたくないこと", value=c.get("not_decide", ""))
            first_check = st.text_area("まず確認したいこと", value=c.get("first_check", ""))
            free_memo = st.text_area("自由メモ", value=c.get("free_memo", ""))
            internal_memo = st.text_area("内部メモ", value=c.get("internal_memo", ""))
            next_check = st.text_area("次回確認すること", value=c.get("next_check", ""))
            submitted = st.form_submit_button("この内容で更新する")
            if submitted:
                update_case_basic(case_id, {
                    "case_title": case_title,
                    "case_type": case_type,
                    "status": status,
                    "consult_date": date_or_blank(consult_date),
                    "next_check_date": date_or_blank(next_check_date),
                    "current_state": current_state,
                    "house_state": house_state,
                    "cat_relation": cat_relation,
                    "family_gap": family_gap,
                    "pressure": pressure,
                    "worries": list_to_text(worries),
                    "not_decide": not_decide,
                    "first_check": first_check,
                    "free_memo": free_memo,
                    "internal_memo": internal_memo,
                    "next_check": next_check,
                })
                st.success("案件を更新しました。")
                st.rerun()

    st.divider()
    st.markdown("### 削除")
    st.warning("削除は元に戻せません。削除前に必ずバックアップしてください。")
    if not df.empty:
        id_cols = {
            "clients": "client_id",
            "cases": "case_id",
            "history": "history_id",
            "properties": "property_id",
            "cats": "cat_id",
            "family": "family_id",
            "photos": "photo_id",
        }
        id_col = id_cols[table]
        delete_id = st.selectbox("削除するID", df[id_col].astype(str).tolist(), key=f"delete_{table}")
        confirm = st.checkbox("本当に削除します", key=f"confirm_delete_{table}")
        if st.button("削除する", type="primary", disabled=not confirm):
            execute(f"DELETE FROM {table} WHERE {id_col}=:id", {"id": delete_id})
            st.success("削除しました。")
            st.rerun()


def page_data_management():
    st.subheader("📦 データ管理")
    st.caption("SQLite版のバックアップ・復元・Excel移行を行います。")

    c1, c2, c3 = st.columns(3)
    c1.metric("相談者", table_count("clients"))
    c2.metric("案件", table_count("cases"))
    c3.metric("履歴", table_count("history"))

    st.markdown("### バックアップ")
    backup_bytes = make_backup_zip_bytes()
    st.download_button(
        "SQLite DB＋写真をZIPバックアップ",
        data=backup_bytes,
        file_name=f"nyantomo_sqlite_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
        mime="application/zip",
    )

    st.download_button(
        "Excel形式でエクスポート",
        data=export_all_to_excel_bytes(),
        file_name=f"nyantomo_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("### 復元")
    st.warning("復元すると現在のSQLite DBが上書きされます。復元前に必ずバックアップしてください。")
    restore_file = st.file_uploader("Ver2.0バックアップZIPをアップロードして復元", type=["zip"], key="restore_zip")
    if restore_file and st.button("バックアップから復元する"):
        restored = restore_backup_zip(restore_file)
        init_db()
        st.success(f"復元しました：{', '.join(restored)}")
        st.rerun()

    st.markdown("### Excel版から移行")
    if OLD_EXCEL_FILE.exists():
        st.info("既存の nyantomo_consultation_data.xlsx が見つかりました。SQLiteへ取り込めます。")
        if st.button("既存ExcelからSQLiteへ移行する"):
            result = import_excel_to_sqlite()
            st.success(f"移行しました：{result}")
            st.rerun()

    excel_file = st.file_uploader("Excel版データをアップロードしてSQLiteへ移行", type=["xlsx"], key="excel_import")
    if excel_file and st.button("アップロードExcelを取り込む"):
        result = import_excel_to_sqlite(excel_file.getvalue())
        st.success(f"移行しました：{result}")
        st.rerun()

    st.markdown("### DBファイル")
    if DB_FILE.exists():
        st.write(f"DBファイル：`{DB_FILE}`")
        st.write(f"サイズ：{DB_FILE.stat().st_size:,} bytes")


# -----------------------------
# 起動
# -----------------------------
init_db()
render_top_nav()

st.title("🐾 にゃんとも相談管理システム Ver2.0（SQLite版）")
st.caption("相談を保留のまま管理する現場OS｜client_id・case_idを正式な主キーとしてDB管理")

# 初回だけExcel移行案内
if not db_has_data() and OLD_EXCEL_FILE.exists():
    with st.container(border=True):
        st.warning("SQLite DBにはまだデータがありません。既存のExcelデータが見つかりました。")
        if st.button("今すぐExcel版データをSQLiteへ移行する", key="top_import_excel"):
            result = import_excel_to_sqlite()
            st.success(f"移行しました：{result}")
            st.rerun()

tabs = st.tabs([
    "🏠 案件ホーム",
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
    "🤖 AI/PDF",
    "🔎 検索・更新・削除",
    "📦 データ管理",
])

with tabs[0]:
    page_case_home()
with tabs[1]:
    page_client_register()
with tabs[2]:
    page_case_register()
with tabs[3]:
    page_case_dashboard()
with tabs[4]:
    page_history()
with tabs[5]:
    page_hold_list()
with tabs[6]:
    page_property()
with tabs[7]:
    page_map()
with tabs[8]:
    page_cat()
with tabs[9]:
    page_family()
with tabs[10]:
    page_photos()
with tabs[11]:
    page_ai_pdf()
with tabs[12]:
    page_search_update_delete()
with tabs[13]:
    page_data_management()
