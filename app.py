
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
import hashlib
import hmac
import secrets
from urllib.parse import quote_plus

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase import pdfmetrics


# =========================================================
# にゃんとも相談管理システム Ver2.5.1 伴走支援版
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
# ログイン・権限管理
# -----------------------------
ROLE_OPTIONS = ["管理者", "編集者", "閲覧者"]
ROLE_PERMISSIONS = {
    "管理者": {"read": True, "write": True, "delete": True, "admin": True, "line": True, "ai": True},
    "編集者": {"read": True, "write": True, "delete": False, "admin": False, "line": True, "ai": True},
    "閲覧者": {"read": True, "write": False, "delete": False, "admin": False, "line": False, "ai": False},
}

def hash_password(password: str) -> str:
    salt = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 120000)
    return f"{salt}${digest.hex()}"

def verify_password(password: str, stored: str) -> bool:
    try:
        salt, digest = stored.split("$", 1)
        new_digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 120000).hex()
        return hmac.compare_digest(new_digest, digest)
    except Exception:
        return False

def ensure_default_admin():
    row = fetch_one("SELECT COUNT(*) AS cnt FROM users")
    if row and int(row["cnt"]) == 0:
        execute("""
            INSERT INTO users(user_id, username, display_name, password_hash, role, is_active, created_at)
            VALUES(:user_id, :username, :display_name, :password_hash, :role, '1', :created_at)
        """, {
            "user_id": make_id("user"),
            "username": "admin",
            "display_name": "管理者",
            "password_hash": hash_password("nyantomo2026"),
            "role": "管理者",
            "created_at": now_text(),
        })

def current_user():
    return st.session_state.get("user")

def has_perm(permission: str) -> bool:
    user = current_user()
    if not user:
        return False
    role = user.get("role", "閲覧者")
    return ROLE_PERMISSIONS.get(role, {}).get(permission, False)

def require_write():
    if not has_perm("write"):
        st.warning("閲覧権限のみのため、この操作はできません。")
        return False
    return True

def require_admin():
    if not has_perm("admin"):
        st.warning("管理者のみ操作できます。")
        return False
    return True

def add_audit_log(action, target_type="", target_id="", detail=""):
    user = current_user() or {}
    try:
        execute("""
            INSERT INTO audit_logs(log_id, user_id, username, action, target_type, target_id, detail, created_at)
            VALUES(:log_id, :user_id, :username, :action, :target_type, :target_id, :detail, :created_at)
        """, {
            "log_id": make_id("log"),
            "user_id": user.get("user_id", ""),
            "username": user.get("username", ""),
            "action": action,
            "target_type": target_type,
            "target_id": target_id,
            "detail": detail,
            "created_at": now_text(),
        })
    except Exception:
        pass

def login_screen():
    st.title("🐾 にゃんとも相談管理システム")
    st.subheader("ログイン")
    st.caption("初期ユーザー：admin ／ 初期パスワード：nyantomo2026")
    with st.form("login_form"):
        username = st.text_input("ユーザー名")
        password = st.text_input("パスワード", type="password")
        submitted = st.form_submit_button("ログイン")
        if submitted:
            user = fetch_one("SELECT * FROM users WHERE username=:username AND is_active='1'", {"username": username})
            if user and verify_password(password, user["password_hash"]):
                st.session_state["user"] = {
                    "user_id": user["user_id"],
                    "username": user["username"],
                    "display_name": user.get("display_name", ""),
                    "role": user["role"],
                }
                add_audit_log("login", "user", user["user_id"], "ログイン")
                st.success("ログインしました。")
                st.rerun()
            else:
                st.error("ユーザー名またはパスワードが違います。")

def logout_button():
    user = current_user()
    if user:
        c1, c2 = st.columns([5, 1])
        with c1:
            st.caption(f"ログイン中：{user.get('display_name') or user.get('username')} ／ 権限：{user.get('role')}")
        with c2:
            if st.button("ログアウト"):
                add_audit_log("logout", "user", user.get("user_id", ""), "ログアウト")
                st.session_state.pop("user", None)
                st.rerun()


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

        CREATE TABLE IF NOT EXISTS users (
            user_id TEXT PRIMARY KEY,
            username TEXT UNIQUE NOT NULL,
            display_name TEXT,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL,
            is_active TEXT DEFAULT '1',
            created_at TEXT
        );

        CREATE TABLE IF NOT EXISTS audit_logs (
            log_id TEXT PRIMARY KEY,
            user_id TEXT,
            username TEXT,
            action TEXT,
            target_type TEXT,
            target_id TEXT,
            detail TEXT,
            created_at TEXT
        );

        CREATE TABLE IF NOT EXISTS ai_summaries (
            summary_id TEXT PRIMARY KEY,
            case_id TEXT NOT NULL,
            client_id TEXT NOT NULL,
            created_at TEXT,
            created_by TEXT,
            summary_type TEXT,
            source_memo TEXT,
            ai_prompt TEXT,
            ai_result TEXT,
            note TEXT,
            FOREIGN KEY(case_id) REFERENCES cases(case_id) ON DELETE CASCADE,
            FOREIGN KEY(client_id) REFERENCES clients(client_id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS line_settings (
            setting_id TEXT PRIMARY KEY,
            channel_access_token TEXT,
            channel_secret TEXT,
            default_to TEXT,
            enabled TEXT DEFAULT '0',
            note TEXT,
            updated_at TEXT
        );

        CREATE TABLE IF NOT EXISTS line_messages (
            message_id TEXT PRIMARY KEY,
            case_id TEXT,
            client_id TEXT,
            created_at TEXT,
            created_by TEXT,
            to_target TEXT,
            message_text TEXT,
            send_status TEXT,
            response_memo TEXT,
            FOREIGN KEY(case_id) REFERENCES cases(case_id) ON DELETE SET NULL,
            FOREIGN KEY(client_id) REFERENCES clients(client_id) ON DELETE SET NULL
        );

        CREATE TABLE IF NOT EXISTS hearing_checklist (
            checklist_id TEXT PRIMARY KEY,
            case_id TEXT NOT NULL,
            client_id TEXT NOT NULL,
            created_at TEXT,
            created_by TEXT,
            item TEXT,
            category TEXT,
            checked TEXT DEFAULT '0',
            checked_at TEXT,
            note TEXT,
            source TEXT,
            FOREIGN KEY(case_id) REFERENCES cases(case_id) ON DELETE CASCADE,
            FOREIGN KEY(client_id) REFERENCES clients(client_id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS state_changes (
            change_id TEXT PRIMARY KEY,
            case_id TEXT NOT NULL,
            client_id TEXT NOT NULL,
            created_at TEXT,
            created_by TEXT,
            field_name TEXT,
            before_value TEXT,
            after_value TEXT,
            reason TEXT,
            memo TEXT,
            FOREIGN KEY(case_id) REFERENCES cases(case_id) ON DELETE CASCADE,
            FOREIGN KEY(client_id) REFERENCES clients(client_id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS follow_suggestions (
            suggestion_id TEXT PRIMARY KEY,
            case_id TEXT NOT NULL,
            client_id TEXT NOT NULL,
            created_at TEXT,
            created_by TEXT,
            suggestion_type TEXT,
            content TEXT,
            priority TEXT,
            status TEXT DEFAULT '未対応',
            due_date TEXT,
            note TEXT,
            FOREIGN KEY(case_id) REFERENCES cases(case_id) ON DELETE CASCADE,
            FOREIGN KEY(client_id) REFERENCES clients(client_id) ON DELETE CASCADE
        );

        """)
        # 将来追加分に備えた軽いマイグレーション
        for col in ["next_check_date", "closed_date", "close_reason", "final_memo", "reopen_possibility", "updated_at"]:
            add_column_if_missing(conn, "cases", col)
        conn.execute("INSERT OR REPLACE INTO meta(key, value) VALUES('app_version', '2.5.1')")
        conn.executescript('''
        CREATE INDEX IF NOT EXISTS idx_cases_client_id ON cases(client_id);
        CREATE INDEX IF NOT EXISTS idx_cases_status ON cases(status);
        CREATE INDEX IF NOT EXISTS idx_cases_next_check_date ON cases(next_check_date);
        CREATE INDEX IF NOT EXISTS idx_history_case_id ON history(case_id);
        CREATE INDEX IF NOT EXISTS idx_properties_case_id ON properties(case_id);
        CREATE INDEX IF NOT EXISTS idx_cats_case_id ON cats(case_id);
        CREATE INDEX IF NOT EXISTS idx_family_case_id ON family(case_id);
        CREATE INDEX IF NOT EXISTS idx_photos_case_id ON photos(case_id);
        CREATE INDEX IF NOT EXISTS idx_ai_summaries_case_id ON ai_summaries(case_id);
        CREATE INDEX IF NOT EXISTS idx_line_messages_case_id ON line_messages(case_id);
        ''')
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
# リレーション整合性チェック・修復
# -----------------------------
def relation_health_check():
    """DB内の親子関係・client_id不一致・写真ファイル有無を確認する。"""
    checks = []

    queries = [
        ("cases_without_client", "案件に対応する相談者が存在しない", """
            SELECT c.case_id AS id, c.client_id, c.case_title AS title
            FROM cases c LEFT JOIN clients cl ON c.client_id = cl.client_id
            WHERE cl.client_id IS NULL
        """),
        ("history_without_case", "履歴に対応する案件が存在しない", """
            SELECT h.history_id AS id, h.case_id, h.client_id, h.record AS title
            FROM history h LEFT JOIN cases c ON h.case_id = c.case_id
            WHERE c.case_id IS NULL
        """),
        ("properties_without_case", "空き家カードに対応する案件が存在しない", """
            SELECT p.property_id AS id, p.case_id, p.client_id, p.property_name AS title
            FROM properties p LEFT JOIN cases c ON p.case_id = c.case_id
            WHERE c.case_id IS NULL
        """),
        ("cats_without_case", "猫情報に対応する案件が存在しない", """
            SELECT ca.cat_id AS id, ca.case_id, ca.client_id, ca.cat_name AS title
            FROM cats ca LEFT JOIN cases c ON ca.case_id = c.case_id
            WHERE c.case_id IS NULL
        """),
        ("family_without_case", "家族メモに対応する案件が存在しない", """
            SELECT f.family_id AS id, f.case_id, f.client_id, f.person_name AS title
            FROM family f LEFT JOIN cases c ON f.case_id = c.case_id
            WHERE c.case_id IS NULL
        """),
        ("photos_without_case", "写真に対応する案件が存在しない", """
            SELECT p.photo_id AS id, p.case_id, p.client_id, p.original_filename AS title
            FROM photos p LEFT JOIN cases c ON p.case_id = c.case_id
            WHERE c.case_id IS NULL
        """),
        ("ai_without_case", "AI履歴に対応する案件が存在しない", """
            SELECT a.summary_id AS id, a.case_id, a.client_id, a.summary_type AS title
            FROM ai_summaries a LEFT JOIN cases c ON a.case_id = c.case_id
            WHERE c.case_id IS NULL
        """),
    ]

    for code, label, sql in queries:
        df = fetch_df(sql)
        if not df.empty:
            df.insert(0, "check", label)
            checks.append(df)

    mismatch_queries = [
        ("history_client_mismatch", "履歴のclient_idが案件のclient_idと不一致", "history", "history_id"),
        ("properties_client_mismatch", "空き家カードのclient_idが案件のclient_idと不一致", "properties", "property_id"),
        ("cats_client_mismatch", "猫情報のclient_idが案件のclient_idと不一致", "cats", "cat_id"),
        ("family_client_mismatch", "家族メモのclient_idが案件のclient_idと不一致", "family", "family_id"),
        ("photos_client_mismatch", "写真のclient_idが案件のclient_idと不一致", "photos", "photo_id"),
        ("ai_client_mismatch", "AI履歴のclient_idが案件のclient_idと不一致", "ai_summaries", "summary_id"),
    ]
    for code_name, label, table, id_col in mismatch_queries:
        df = fetch_df(f"""
            SELECT t.{id_col} AS id, t.case_id, t.client_id AS child_client_id, c.client_id AS case_client_id
            FROM {table} t JOIN cases c ON t.case_id = c.case_id
            WHERE COALESCE(t.client_id, '') != COALESCE(c.client_id, '')
        """)
        if not df.empty:
            df.insert(0, "check", label)
            checks.append(df)

    photo_df = fetch_df("SELECT photo_id AS id, case_id, client_id, saved_path FROM photos WHERE saved_path IS NOT NULL AND saved_path != ''")
    missing_photos = []
    for _, r in photo_df.iterrows():
        if not Path(str(r.get("saved_path", ""))).exists():
            missing_photos.append(dict(r))
    if missing_photos:
        df = pd.DataFrame(missing_photos)
        df.insert(0, "check", "写真DBはあるがファイルが存在しない")
        checks.append(df)

    if checks:
        return pd.concat(checks, ignore_index=True, sort=False)
    return pd.DataFrame(columns=["check", "id", "case_id", "client_id", "title"])


def repair_child_client_ids():
    """case_idを正として、子テーブルのclient_idを案件側に揃える。"""
    updated = {}
    targets = [
        ("history", "history_id"),
        ("properties", "property_id"),
        ("cats", "cat_id"),
        ("family", "family_id"),
        ("photos", "photo_id"),
        ("ai_summaries", "summary_id"),
        ("line_messages", "message_id"),
    ]
    with get_conn() as conn:
        for table, id_col in targets:
            cur = conn.execute(f"""
                UPDATE {table}
                SET client_id = (SELECT client_id FROM cases WHERE cases.case_id = {table}.case_id)
                WHERE case_id IS NOT NULL
                  AND EXISTS (SELECT 1 FROM cases WHERE cases.case_id = {table}.case_id)
                  AND COALESCE(client_id, '') != COALESCE((SELECT client_id FROM cases WHERE cases.case_id = {table}.case_id), '')
            """)
            updated[table] = cur.rowcount
        conn.commit()
    add_audit_log("repair_relations", "database", "sqlite", str(updated))
    return updated


def delete_photo_files_for_case(case_id):
    photo_df = fetch_df("SELECT saved_path FROM photos WHERE case_id=:case_id", {"case_id": case_id})
    removed = 0
    for _, r in photo_df.iterrows():
        path = Path(str(r.get("saved_path", "")))
        if path.exists() and path.is_file():
            try:
                path.unlink()
                removed += 1
            except Exception:
                pass
    return removed


def delete_photo_file_by_id(photo_id):
    row = fetch_one("SELECT saved_path FROM photos WHERE photo_id=:photo_id", {"photo_id": photo_id})
    if row:
        path = Path(str(row.get("saved_path", "")))
        if path.exists() and path.is_file():
            try:
                path.unlink()
                return 1
            except Exception:
                return 0
    return 0

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
            "version": "2.2.1",
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
        # 復元時は古い写真ファイルの取り残しを避けるため、photosを入れ替える
        if PHOTO_DIR.exists():
            shutil.rmtree(PHOTO_DIR, ignore_errors=True)
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
# Ver2.5.1 安定稼働：インデックス・整合性チェック
# -----------------------------
def ensure_indexes():
    """検索・関連データ取得を安定化するためのインデックスを作成する。"""
    with get_conn() as conn:
        conn.executescript("""
        CREATE INDEX IF NOT EXISTS idx_cases_client_id ON cases(client_id);
        CREATE INDEX IF NOT EXISTS idx_cases_status ON cases(status);
        CREATE INDEX IF NOT EXISTS idx_cases_next_check_date ON cases(next_check_date);
        CREATE INDEX IF NOT EXISTS idx_cases_updated_at ON cases(updated_at);

        CREATE INDEX IF NOT EXISTS idx_history_case_id ON history(case_id);
        CREATE INDEX IF NOT EXISTS idx_history_client_id ON history(client_id);
        CREATE INDEX IF NOT EXISTS idx_properties_case_id ON properties(case_id);
        CREATE INDEX IF NOT EXISTS idx_properties_client_id ON properties(client_id);
        CREATE INDEX IF NOT EXISTS idx_cats_case_id ON cats(case_id);
        CREATE INDEX IF NOT EXISTS idx_cats_client_id ON cats(client_id);
        CREATE INDEX IF NOT EXISTS idx_family_case_id ON family(case_id);
        CREATE INDEX IF NOT EXISTS idx_family_client_id ON family(client_id);
        CREATE INDEX IF NOT EXISTS idx_photos_case_id ON photos(case_id);
        CREATE INDEX IF NOT EXISTS idx_photos_client_id ON photos(client_id);
        CREATE INDEX IF NOT EXISTS idx_ai_summaries_case_id ON ai_summaries(case_id);
        CREATE INDEX IF NOT EXISTS idx_ai_summaries_client_id ON ai_summaries(client_id);
        CREATE INDEX IF NOT EXISTS idx_line_messages_case_id ON line_messages(case_id);
        CREATE INDEX IF NOT EXISTS idx_line_messages_client_id ON line_messages(client_id);
        CREATE INDEX IF NOT EXISTS idx_hearing_checklist_case_id ON hearing_checklist(case_id);
        CREATE INDEX IF NOT EXISTS idx_state_changes_case_id ON state_changes(case_id);
        CREATE INDEX IF NOT EXISTS idx_follow_suggestions_case_id ON follow_suggestions(case_id);
        """)
        conn.commit()


RELATED_TABLES = [
    ("history", "history_id"),
    ("properties", "property_id"),
    ("cats", "cat_id"),
    ("family", "family_id"),
    ("photos", "photo_id"),
    ("ai_summaries", "summary_id"),
    ("line_messages", "message_id"),
]


def relation_check_df():
    """case_id / client_id のズレ、孤立データ、写真ファイル欠損を一覧化する。"""
    rows = []

    for table, id_col in RELATED_TABLES:
        # case_id が存在しない孤立データ
        orphan_case = fetch_df(f"""
            SELECT r.{id_col} AS record_id, r.case_id, r.client_id
            FROM {table} r
            LEFT JOIN cases c ON r.case_id = c.case_id
            WHERE r.case_id IS NOT NULL
              AND r.case_id != ''
              AND c.case_id IS NULL
        """)
        for _, r in orphan_case.iterrows():
            rows.append({
                "種別": "孤立データ",
                "テーブル": table,
                "ID": r.get("record_id", ""),
                "case_id": r.get("case_id", ""),
                "client_id": r.get("client_id", ""),
                "内容": "親案件が存在しません",
                "自動修復": "不可：削除または案件復元が必要",
            })

        # case_id の親案件 client_id と子データ client_id の不一致
        mismatch = fetch_df(f"""
            SELECT r.{id_col} AS record_id, r.case_id, r.client_id AS child_client_id, c.client_id AS parent_client_id
            FROM {table} r
            JOIN cases c ON r.case_id = c.case_id
            WHERE COALESCE(r.client_id, '') != COALESCE(c.client_id, '')
        """)
        for _, r in mismatch.iterrows():
            rows.append({
                "種別": "client_id不一致",
                "テーブル": table,
                "ID": r.get("record_id", ""),
                "case_id": r.get("case_id", ""),
                "client_id": r.get("child_client_id", ""),
                "内容": f"親案件のclient_id={r.get('parent_client_id', '')} と不一致",
                "自動修復": "可：親案件のclient_idに合わせる",
            })

    # 写真DBにあるがファイルがない
    try:
        photos = fetch_df("SELECT photo_id, case_id, client_id, saved_path FROM photos")
        for _, p in photos.iterrows():
            sp = str(p.get("saved_path", ""))
            if sp and not Path(sp).exists():
                rows.append({
                    "種別": "写真ファイル欠損",
                    "テーブル": "photos",
                    "ID": p.get("photo_id", ""),
                    "case_id": p.get("case_id", ""),
                    "client_id": p.get("client_id", ""),
                    "内容": f"保存先ファイルが見つかりません：{sp}",
                    "自動修復": "不可：再アップロードまたは写真データ削除",
                })
    except Exception:
        pass

    return pd.DataFrame(rows)


def repair_client_id_mismatches():
    """子テーブルのclient_idを親案件のclient_idへ揃える。"""
    total = 0
    with get_conn() as conn:
        for table, id_col in RELATED_TABLES:
            cur = conn.execute(f"""
                UPDATE {table}
                SET client_id = (
                    SELECT cases.client_id
                    FROM cases
                    WHERE cases.case_id = {table}.case_id
                )
                WHERE EXISTS (
                    SELECT 1
                    FROM cases
                    WHERE cases.case_id = {table}.case_id
                      AND COALESCE(cases.client_id, '') != COALESCE({table}.client_id, '')
                )
            """)
            total += cur.rowcount if cur.rowcount is not None else 0
        conn.commit()
    add_audit_log("repair_relations", "database", "client_id_mismatch", f"fixed={total}")
    return total


def delete_orphan_related_records():
    """親案件がない孤立データを削除する。通常は使わないが、復元失敗時の整理用。"""
    total = 0
    with get_conn() as conn:
        for table, id_col in RELATED_TABLES:
            cur = conn.execute(f"""
                DELETE FROM {table}
                WHERE case_id IS NOT NULL
                  AND case_id != ''
                  AND NOT EXISTS (
                      SELECT 1 FROM cases WHERE cases.case_id = {table}.case_id
                  )
            """)
            total += cur.rowcount if cur.rowcount is not None else 0
        conn.commit()
    add_audit_log("delete_orphans", "database", "orphan_related_records", f"deleted={total}")
    return total


def delete_photo_files_for_case(case_id):
    """案件削除前に紐づく写真ファイルを削除する。"""
    removed = 0
    try:
        df = fetch_df("SELECT saved_path FROM photos WHERE case_id=:case_id", {"case_id": case_id})
        for _, row in df.iterrows():
            path = Path(str(row.get("saved_path", "")))
            if path.exists() and path.is_file():
                try:
                    path.unlink()
                    removed += 1
                except Exception:
                    pass
    except Exception:
        pass
    return removed


def delete_photo_file_by_id(photo_id):
    """写真データ削除前に実ファイルを削除する。"""
    removed = 0
    try:
        row = fetch_one("SELECT saved_path FROM photos WHERE photo_id=:photo_id", {"photo_id": photo_id})
        if row:
            path = Path(str(row.get("saved_path", "")))
            if path.exists() and path.is_file():
                try:
                    path.unlink()
                    removed += 1
                except Exception:
                    pass
    except Exception:
        pass
    return removed


def touch_case(case_id):
    """子データ更新時に親案件の最終更新日を更新する。"""
    if case_id:
        execute("UPDATE cases SET updated_at=:updated_at WHERE case_id=:case_id", {
            "updated_at": now_text(),
            "case_id": case_id,
        })


def safe_restore_backup_zip(uploaded_file):
    """
    復元前に現行DBを退避し、復元後にDB初期化・インデックス作成を行う。
    """
    safety_dir = Path("_safety_backup_before_restore")
    safety_dir.mkdir(exist_ok=True)
    if DB_FILE.exists():
        shutil.copy2(DB_FILE, safety_dir / f"{DB_FILE.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db")

    # 既存photosを一度退避。復元ZIPにphotosがある場合は上書き整理。
    if PHOTO_DIR.exists():
        backup_photo_dir = safety_dir / f"photos_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        try:
            shutil.copytree(PHOTO_DIR, backup_photo_dir, dirs_exist_ok=True)
        except Exception:
            pass

    restored = restore_backup_zip(uploaded_file)
    init_db()
    ensure_indexes()
    add_audit_log("restore_backup", "database", "backup_zip", ",".join(restored))
    return restored


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



# -----------------------------
# Ver2.5.1 伴走支援
# -----------------------------
def get_table_count_by_case(table, case_id):
    try:
        row = fetch_one(f"SELECT COUNT(*) AS cnt FROM {table} WHERE case_id=:case_id", {"case_id": case_id})
        return int(row.get("cnt", 0)) if row else 0
    except Exception:
        return 0


def generate_hearing_items(case_id):
    c = get_case_full(case_id)
    if not c:
        return []

    items = []
    status = c.get("status", "")
    house_state = c.get("house_state", "")
    cat_relation = c.get("cat_relation", "")
    family_gap = c.get("family_gap", "")
    pressure = c.get("pressure", "")
    worries = c.get("worries", "")
    not_decide = c.get("not_decide", "")
    first_check = c.get("first_check", "")
    next_check = c.get("next_check", "")

    def add(category, item):
        if item not in [x["item"] for x in items]:
            items.append({"category": category, "item": item})

    add("基本確認", "前回から状況や気持ちに変化があったか。")
    add("基本確認", "今も『急いで決めたくないこと』は同じか、変わったか。")
    add("基本確認", "次回までに確認だけならできそうなことは何か。")

    if house_state in ["未選択", ""]:
        add("住まい", "現在の住まいの状態を確認する。")
    if house_state in ["空き家になっている", "近いうちに空き家になりそう", "相続後そのまま", "売却・賃貸を迷っている"]:
        add("住まい", "郵便物・通風・雨漏り・庭・近隣不安など、空き家の状態に変化がないか。")
        add("住まい", "鍵の所在、現地確認できる人、緊急時の連絡先を確認する。")
    if get_table_count_by_case("properties", case_id) == 0 and ("空き家" in worries or house_state in ["空き家になっている", "近いうちに空き家になりそう", "相続後そのまま"]):
        add("住まい", "空き家カードに登録するため、所在地・状態・管理頻度を確認する。")

    if cat_relation in ["未選択", ""]:
        add("猫", "猫との関係、現在の世話の体制、今後心配していることを確認する。")
    if cat_relation in ["猫と暮らしている", "家族の猫がいる", "猫を残して入院・施設入所が心配", "これから猫と暮らしたい"] or "猫" in worries:
        add("猫", "猫の現在の暮らし、世話をしている人、緊急時の預け先候補を確認する。")
        add("猫", "猫について『今は決めたくないこと』があるか確認する。")
    if get_table_count_by_case("cats", case_id) == 0 and ("猫" in cat_relation or "猫" in worries):
        add("猫", "猫情報カードに登録するため、名前・頭数・年齢・現在の暮らしを確認する。")

    if family_gap in ["未選択", "", "まだ話せていない"]:
        add("家族", "家族とどこまで話せているか、まだ話せていない相手がいるか確認する。")
    if family_gap in ["少しある", "かなりある", "まだ話せていない"]:
        add("家族", "誰が急いでいて、誰が迷っているのかを対立させずに確認する。")
        add("家族", "相談者本人の希望と家族の希望がどこで違うのか確認する。")
    if get_table_count_by_case("family", case_id) == 0 and family_gap in ["少しある", "かなりある", "まだ話せていない"]:
        add("家族", "家族関係メモに登録するため、関係者名・続柄・温度感を確認する。")

    if pressure in ["少しある", "強くある", "自分でも焦っている"]:
        add("急がされ感", "何に急かされていると感じているか。期限・家族・お金・空き家劣化などを分ける。")
        add("急がされ感", "本当に急ぐ必要があることと、急がなくてもよいことを分ける。")

    if not not_decide:
        add("保留", "今は決めたくないことを確認する。")
    if not first_check:
        add("保留", "まず確認だけしたいことを確認する。")
    if not next_check:
        add("次回", "次回までの小さな確認事項を1〜3個に絞る。")

    return items


def generate_missing_warnings(case_id):
    c = get_case_full(case_id)
    if not c:
        return []
    warnings = []

    def warn(level, item, detail):
        warnings.append({"level": level, "item": item, "detail": detail})

    for key, label in [
        ("current_state", "今いちばん近い状態"),
        ("house_state", "住まいの状態"),
        ("cat_relation", "猫との関係"),
        ("family_gap", "家族との温度差"),
        ("pressure", "急がされている感じ"),
        ("worries", "気になること"),
    ]:
        v = c.get(key, "")
        if not v or v == "未選択":
            warn("確認", label, "未確認です。次回ヒアリングで確認すると安全です。")

    if not c.get("next_check_date"):
        warn("重要", "次回確認日", "未設定です。保留を放置にしないため、確認日を設定してください。")
    if not c.get("next_check"):
        warn("確認", "次回確認すること", "未設定です。次回の小さな確認事項を残してください。")

    last_days = get_last_update_days(c.get("updated_at", ""))
    if last_days is not None and last_days >= 60:
        warn("重要", "未更新", f"最終更新から{last_days}日経過しています。急かさない形で状況確認を検討してください。")

    house_state = c.get("house_state", "")
    if house_state in ["空き家になっている", "近いうちに空き家になりそう", "相続後そのまま"] and get_table_count_by_case("properties", case_id) == 0:
        warn("重要", "空き家カード", "空き家状態に近いですが、空き家カードが未登録です。")
    if ("猫" in c.get("cat_relation", "") or "猫" in c.get("worries", "")) and get_table_count_by_case("cats", case_id) == 0:
        warn("確認", "猫情報カード", "猫に関する相談要素がありますが、猫情報カードが未登録です。")
    if c.get("family_gap", "") in ["少しある", "かなりある", "まだ話せていない"] and get_table_count_by_case("family", case_id) == 0:
        warn("確認", "家族関係メモ", "家族温度差がありますが、家族関係メモが未登録です。")

    return warnings


def generate_do_not_do(case_id):
    c = get_case_full(case_id)
    if not c:
        return []
    items = []
    pressure = c.get("pressure", "")
    family_gap = c.get("family_gap", "")
    house_state = c.get("house_state", "")
    cat_relation = c.get("cat_relation", "")
    worries = c.get("worries", "")

    items.append("売却・賃貸・施設入所などの結論をこちらから急がせない。")
    items.append("『正解』を提示するより、事実・未確定・保留を分ける。")

    if pressure in ["少しある", "強くある", "自分でも焦っている"]:
        items.append("焦りがある状態で大きな判断を迫らない。")
    if family_gap in ["少しある", "かなりある", "まだ話せていない"]:
        items.append("家族会議や家族間調整を急がせない。まず温度差を言葉にする。")
    if "猫" in cat_relation or "猫" in worries:
        items.append("猫の譲渡・別居・預け先の結論を急がない。まず緊急時の選択肢を確認する。")
    if house_state in ["空き家になっている", "近いうちに空き家になりそう", "相続後そのまま"]:
        items.append("空き家をすぐ売る・貸す方向に寄せない。まず安全確認と管理状態を把握する。")

    return list(dict.fromkeys(items))


def generate_follow_suggestions(case_id):
    c = get_case_full(case_id)
    if not c:
        return []
    suggestions = []

    def add(kind, content, priority="確認", days=14):
        due = (date.today() + pd.Timedelta(days=days)).strftime("%Y-%m-%d")
        suggestions.append({"type": kind, "content": content, "priority": priority, "due_date": due})

    add("短期確認", "次回確認日までに、状況変化がないか短い連絡で確認する。", "確認", 14)

    if c.get("pressure") in ["少しある", "強くある", "自分でも焦っている"]:
        add("安心形成", "急ぐ話と急がなくてよい話を分けるため、短時間の整理面談を提案する。", "重要", 7)
    if c.get("house_state") in ["空き家になっている", "近いうちに空き家になりそう", "相続後そのまま"]:
        add("住まい確認", "空き家の通風・郵便物・外観・近隣不安の確認を提案する。", "重要", 14)
    if "猫" in c.get("cat_relation", "") or "猫" in c.get("worries", ""):
        add("猫確認", "猫の世話体制・緊急時の預け先候補を確認する。", "確認", 14)
    if c.get("family_gap") in ["少しある", "かなりある", "まだ話せていない"]:
        add("家族温度差", "家族の誰が何を心配しているのか、対立させずに整理する機会を作る。", "確認", 21)

    return suggestions


def save_generated_hearing_items(case_id, items):
    c = get_case_full(case_id)
    if not c:
        return 0
    count = 0
    for it in items:
        exists = fetch_one("""
            SELECT checklist_id FROM hearing_checklist
            WHERE case_id=:case_id AND item=:item
            LIMIT 1
        """, {"case_id": case_id, "item": it["item"]})
        if exists:
            continue
        execute("""
            INSERT INTO hearing_checklist(checklist_id, case_id, client_id, created_at, created_by,
                                          item, category, checked, checked_at, note, source)
            VALUES(:checklist_id, :case_id, :client_id, :created_at, :created_by,
                   :item, :category, '0', '', '', 'auto')
        """, {
            "checklist_id": make_id("hear"),
            "case_id": case_id,
            "client_id": c["client_id"],
            "created_at": now_text(),
            "created_by": (current_user() or {}).get("username", ""),
            "item": it["item"],
            "category": it["category"],
        })
        count += 1
    if count:
        add_audit_log("save_hearing_items", "case", case_id, f"count={count}")
    return count


def save_follow_suggestions(case_id, suggestions):
    c = get_case_full(case_id)
    if not c:
        return 0
    count = 0
    for s in suggestions:
        exists = fetch_one("""
            SELECT suggestion_id FROM follow_suggestions
            WHERE case_id=:case_id AND content=:content
            LIMIT 1
        """, {"case_id": case_id, "content": s["content"]})
        if exists:
            continue
        execute("""
            INSERT INTO follow_suggestions(suggestion_id, case_id, client_id, created_at, created_by,
                                           suggestion_type, content, priority, status, due_date, note)
            VALUES(:suggestion_id, :case_id, :client_id, :created_at, :created_by,
                   :suggestion_type, :content, :priority, '未対応', :due_date, '')
        """, {
            "suggestion_id": make_id("follow"),
            "case_id": case_id,
            "client_id": c["client_id"],
            "created_at": now_text(),
            "created_by": (current_user() or {}).get("username", ""),
            "suggestion_type": s["type"],
            "content": s["content"],
            "priority": s["priority"],
            "due_date": s["due_date"],
        })
        count += 1
    if count:
        add_audit_log("save_follow_suggestions", "case", case_id, f"count={count}")
    return count


def render_companion_support(case_id):
    c = get_case_full(case_id)
    if not c:
        st.error("案件が見つかりません。")
        return

    hearing_items = generate_hearing_items(case_id)
    warnings = generate_missing_warnings(case_id)
    dont_do = generate_do_not_do(case_id)
    follow = generate_follow_suggestions(case_id)

    st.caption("判断を急がせず、次回確認・抜け漏れ・状態変化・継続フォローを支える画面です。")

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("次回ヒアリング案", len(hearing_items))
    k2.metric("漏れ警告", len(warnings))
    k3.metric("今やらないこと", len(dont_do))
    k4.metric("フォロー提案", len(follow))

    support_tabs = st.tabs(["次回ヒアリング", "漏れ警告", "状態変化ログ", "今やらないこと", "継続フォロー"])

    with support_tabs[0]:
        st.markdown("### 次回ヒアリング項目")
        if hearing_items:
            df = pd.DataFrame(hearing_items)
            st.dataframe(df, use_container_width=True)
            for idx, it in enumerate(hearing_items):
                st.checkbox(
                    f"{it['category']}｜{it['item']}",
                    key=f"preview_hear_{case_id}_{idx}",
                    disabled=True
                )
            if st.button("このヒアリング項目を保存する", key=f"save_hearing_{case_id}", disabled=not has_perm("write")):
                cnt = save_generated_hearing_items(case_id, hearing_items)
                st.success(f"{cnt}件保存しました。")
                st.rerun()
        else:
            st.success("追加で提案するヒアリング項目は少ないです。")

        st.markdown("### 保存済みチェックリスト")
        saved = fetch_df("""
            SELECT checklist_id, created_at, category, item, checked, checked_at, note
            FROM hearing_checklist
            WHERE case_id=:case_id
            ORDER BY created_at DESC
        """, {"case_id": case_id})
        st.dataframe(saved, use_container_width=True)

        if not saved.empty and has_perm("write"):
            selected = st.selectbox(
                "チェック更新する項目",
                [f"{r['category']}｜{r['item']}｜{r['checklist_id']}" for _, r in saved.iterrows()],
                key=f"checklist_update_select_{case_id}"
            )
            checklist_id = selected_id_from_label(selected)
            row = fetch_one("SELECT * FROM hearing_checklist WHERE checklist_id=:id", {"id": checklist_id})
            with st.form(f"checklist_update_form_{checklist_id}"):
                checked = st.selectbox("確認状態", ["0", "1"], index=1 if row.get("checked") == "1" else 0, format_func=lambda x: "確認済" if x == "1" else "未確認")
                note = st.text_area("確認メモ", value=row.get("note", ""))
                submitted = st.form_submit_button("チェックリストを更新")
                if submitted:
                    execute("""
                        UPDATE hearing_checklist
                        SET checked=:checked, checked_at=:checked_at, note=:note
                        WHERE checklist_id=:id
                    """, {
                        "checked": checked,
                        "checked_at": now_text() if checked == "1" else "",
                        "note": note,
                        "id": checklist_id,
                    })
                    touch_case(case_id)
                    add_audit_log("update_hearing_checklist", "case", case_id, checklist_id)
                    st.success("更新しました。")
                    st.rerun()

    with support_tabs[1]:
        st.markdown("### ヒアリング漏れ警告")
        if warnings:
            wdf = pd.DataFrame(warnings)
            st.dataframe(wdf, use_container_width=True)
            for w in warnings:
                if w["level"] == "重要":
                    st.warning(f"⚠ {w['item']}：{w['detail']}")
                else:
                    st.info(f"確認：{w['item']}：{w['detail']}")
        else:
            st.success("大きなヒアリング漏れは見つかりません。")

    with support_tabs[2]:
        st.markdown("### 状態変化ログ")
        state_df = fetch_df("""
            SELECT created_at, created_by, field_name, before_value, after_value, reason, memo, change_id
            FROM state_changes
            WHERE case_id=:case_id
            ORDER BY created_at DESC
        """, {"case_id": case_id})
        st.dataframe(state_df, use_container_width=True)

        if has_perm("write"):
            with st.form(f"state_change_form_{case_id}"):
                field_name = st.selectbox("変化した項目", ["急がされ感", "家族温度差", "本人の不安", "住まいの状態", "猫の状況", "ステータス", "その他"])
                before_value = st.text_input("変化前")
                after_value = st.text_input("変化後")
                reason = st.text_input("変化のきっかけ")
                memo = st.text_area("メモ")
                submitted = st.form_submit_button("状態変化を記録")
                if submitted:
                    execute("""
                        INSERT INTO state_changes(change_id, case_id, client_id, created_at, created_by,
                                                  field_name, before_value, after_value, reason, memo)
                        VALUES(:change_id, :case_id, :client_id, :created_at, :created_by,
                               :field_name, :before_value, :after_value, :reason, :memo)
                    """, {
                        "change_id": make_id("chg"),
                        "case_id": case_id,
                        "client_id": c["client_id"],
                        "created_at": now_text(),
                        "created_by": (current_user() or {}).get("username", ""),
                        "field_name": field_name,
                        "before_value": before_value,
                        "after_value": after_value,
                        "reason": reason,
                        "memo": memo,
                    })
                    touch_case(case_id)
                    add_audit_log("add_state_change", "case", case_id, field_name)
                    st.success("状態変化を記録しました。")
                    st.rerun()

        st.markdown("### ステータス履歴")
        hist = fetch_df("""
            SELECT record_date, record_type, before_status, after_status, record
            FROM history
            WHERE case_id=:case_id
            ORDER BY record_date DESC, created_at DESC
        """, {"case_id": case_id})
        st.dataframe(hist, use_container_width=True)

    with support_tabs[3]:
        st.markdown("### 今やらない方がいいこと")
        for item in dont_do:
            st.warning(f"・{item}")
        st.caption("これは禁止ではなく、相談者の判断を奪わないための安全装置です。")

    with support_tabs[4]:
        st.markdown("### 継続フォロー提案")
        fdf = pd.DataFrame(follow)
        st.dataframe(fdf, use_container_width=True)
        if st.button("このフォロー提案を保存する", key=f"save_follow_{case_id}", disabled=not has_perm("write")):
            cnt = save_follow_suggestions(case_id, follow)
            st.success(f"{cnt}件保存しました。")
            st.rerun()

        st.markdown("### 保存済みフォロー")
        saved_follow = fetch_df("""
            SELECT suggestion_id, created_at, suggestion_type, priority, status, due_date, content, note
            FROM follow_suggestions
            WHERE case_id=:case_id
            ORDER BY due_date ASC, created_at DESC
        """, {"case_id": case_id})
        st.dataframe(saved_follow, use_container_width=True)

        if not saved_follow.empty and has_perm("write"):
            selected = st.selectbox(
                "更新するフォロー",
                [f"{r['due_date']}｜{r['status']}｜{r['content'][:30]}｜{r['suggestion_id']}" for _, r in saved_follow.iterrows()],
                key=f"follow_update_select_{case_id}"
            )
            suggestion_id = selected_id_from_label(selected)
            row = fetch_one("SELECT * FROM follow_suggestions WHERE suggestion_id=:id", {"id": suggestion_id})
            with st.form(f"follow_update_form_{suggestion_id}"):
                status = st.selectbox("状態", ["未対応", "確認中", "対応済", "保留", "不要"], index=option_index(["未対応", "確認中", "対応済", "保留", "不要"], row.get("status", "")))
                due_date = st.date_input("期限", value=parse_date_safe(row.get("due_date")), key=f"follow_due_{suggestion_id}")
                note = st.text_area("メモ", value=row.get("note", ""))
                submitted = st.form_submit_button("フォローを更新")
                if submitted:
                    execute("""
                        UPDATE follow_suggestions
                        SET status=:status, due_date=:due_date, note=:note
                        WHERE suggestion_id=:id
                    """, {
                        "status": status,
                        "due_date": date_or_blank(due_date),
                        "note": note,
                        "id": suggestion_id,
                    })
                    touch_case(case_id)
                    add_audit_log("update_follow_suggestion", "case", case_id, suggestion_id)
                    st.success("更新しました。")
                    st.rerun()


def page_companion_support():
    st.subheader("🤝 伴走支援")
    case_id = select_case_widget("companion_case_select", include_closed=True)
    if not case_id:
        return
    render_companion_support(case_id)



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

    line_count_row = fetch_one("SELECT COUNT(*) AS cnt FROM line_messages WHERE case_id=:case_id", {"case_id": case_id})
    latest_line = fetch_one("SELECT created_at, send_status FROM line_messages WHERE case_id=:case_id ORDER BY created_at DESC LIMIT 1", {"case_id": case_id})
    if latest_line:
        st.info(f"LINE履歴：{line_count_row.get('cnt', 0)}件 ／ 最新：{latest_line.get('created_at','')}（{latest_line.get('send_status','')}）")
    else:
        st.info("LINE履歴：未登録")

    st.divider()
    st.markdown("### この案件に紐づくデータ")
    rel_tabs = st.tabs(["相談者", "相談履歴", "空き家", "猫", "家族", "写真", "LINE履歴", "伴走支援", "AI/PDF用メモ", "管理者アドバイス"])

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
        line_df = fetch_df("""
            SELECT
                created_at AS 送信時,
                created_by AS 作成者,
                to_target AS 送信先ID,
                send_status AS 送信状態,
                message_text AS 送信文,
                response_memo AS 反応メモ,
                message_id
            FROM line_messages
            WHERE case_id=:case_id
            ORDER BY created_at DESC
        """, {"case_id": case_id})

        if line_df.empty:
            st.info("この案件のLINE送信履歴はまだありません。")
        else:
            st.dataframe(line_df, use_container_width=True)
            selected_line = st.selectbox(
                "送信文を確認する履歴",
                [
                    f"{r['送信時']}｜{r['送信状態']}｜{r['message_id']}"
                    for _, r in line_df.iterrows()
                ],
                key=f"dash_line_history_select_{case_id}"
            )
            message_id = selected_id_from_label(selected_line)
            line_row = fetch_one("SELECT * FROM line_messages WHERE message_id=:message_id", {"message_id": message_id})
            if line_row:
                st.markdown("#### 送信内容")
                st.write(f"**送信時：** {line_row.get('created_at','')}")
                st.write(f"**送信状態：** {line_row.get('send_status','')}")
                st.write(f"**送信先ID：** {line_row.get('to_target','')}")
                st.text_area("送信文", line_row.get("message_text", ""), height=220, key=f"dash_line_text_{message_id}")
                st.text_area("反応メモ", line_row.get("response_memo", ""), height=120, key=f"dash_line_response_{message_id}")

    with rel_tabs[7]:
        render_companion_support(case_id)

    with rel_tabs[8]:
        memo = build_case_memo(case_id)
        st.text_area("案件統合メモ", memo, height=500)
        st.download_button("この案件のPDFをダウンロード", make_pdf_bytes(memo), file_name=f"nyantomo_case_{case_id}.pdf", mime="application/pdf")

    with rel_tabs[9]:
        advice = build_management_advice(case_id)
        st.text_area("管理者向け対応アドバイス案", advice, height=500)
        st.download_button("対応アドバイスPDFをダウンロード", make_pdf_bytes(advice), file_name=f"nyantomo_advice_{case_id}.pdf", mime="application/pdf")


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



def build_management_advice(case_id):
    """
    案件ごとの現時点情報から、管理者向けの対応アドバイス案を作成する。
    注意：これは判断代行ではなく、事実・未確定・保留・次回確認を整理する補助出力。
    """
    c = get_case_full(case_id)
    if not c:
        return ""

    history_df = get_related_df("history", case_id)
    prop_df = get_related_df("properties", case_id)
    cat_df = get_related_df("cats", case_id)
    fam_df = get_related_df("family", case_id)
    photo_df = get_related_df("photos", case_id)

    status = c.get("status", "")
    pressure = c.get("pressure", "")
    family_gap = c.get("family_gap", "")
    house_state = c.get("house_state", "")
    cat_relation = c.get("cat_relation", "")
    worries = c.get("worries", "")
    next_date = c.get("next_check_date", "")
    last_update_days = get_last_update_days(c.get("updated_at", ""))

    # リスク・確認ポイントを軽く判定
    attention = []
    if pressure in ["強くある", "自分でも焦っている"]:
        attention.append("相談者が急がされている／焦っている可能性があるため、結論よりも状況整理を優先する。")
    if family_gap in ["かなりある", "まだ話せていない"]:
        attention.append("家族間の温度差があるため、対立構造にせず、事実と気持ちを分けて確認する。")
    if house_state in ["空き家になっている", "近いうちに空き家になりそう", "相続後そのまま"]:
        attention.append("住まいが空白状態に近いため、売却・賃貸判断ではなく、まず管理・安全確認の必要性を整理する。")
    if cat_relation in ["猫を残して入院・施設入所が心配", "猫と暮らしている", "家族の猫がいる"]:
        attention.append("猫の生活継続が判断に影響するため、人・住まい・猫を分けずに確認する。")
    if last_update_days is not None and last_update_days >= 60:
        attention.append(f"最終更新から{last_update_days}日経過しているため、急かさない形で状況確認を入れる。")
    if not next_date:
        attention.append("次回確認日が未設定のため、保留と放置を分けるために確認日を設定する。")

    if not attention:
        attention.append("現時点で強い注意点は少ないが、未確定事項を分けて静かに整理する。")

    # 不足情報
    missing = []
    missing_map = {
        "current_state": "相談者が今どの段階にいるか",
        "house_state": "住まいの状態",
        "cat_relation": "猫との関係",
        "family_gap": "家族との温度差",
        "pressure": "急がされている感じ",
        "worries": "気になること",
        "next_check_date": "次回確認日",
    }
    for key, label in missing_map.items():
        v = c.get(key, "")
        if not v or v == "未選択":
            missing.append(label)

    # 次回の具体案
    next_actions = []
    if "空き家管理" in worries or house_state in ["空き家になっている", "近いうちに空き家になりそう", "相続後そのまま"]:
        next_actions.append("住まいについて、現在の管理状況・鍵の所在・郵便物・近隣不安の有無を確認する。")
    if "猫" in cat_relation or "猫" in worries or not cat_df.empty:
        next_actions.append("猫について、現在の世話の体制・緊急時の預け先候補・本人の希望を確認する。")
    if family_gap in ["少しある", "かなりある", "まだ話せていない"]:
        next_actions.append("家族について、誰が急いでいるのか／誰が迷っているのかを、責めずに分けて聞く。")
    if not next_actions:
        next_actions.append("次回は、相談者が『今は決めたくないこと』と『確認だけならできること』を分けて聞く。")

    # 管理者向け声かけ案
    talk = []
    talk.append("「今日すぐに決めなくて大丈夫です。まず、今わかっていることと、まだ決めなくてよいことを分けてみましょう。」")
    if pressure in ["少しある", "強くある", "自分でも焦っている"]:
        talk.append("「急いだ方がよい話と、急がなくてもよい話が混ざっているかもしれません。一度、分けて確認しましょう。」")
    if family_gap in ["少しある", "かなりある", "まだ話せていない"]:
        talk.append("「ご家族の意見も大切ですが、まずはご本人が何に困っているかを確認してから進めましょう。」")
    if "猫" in cat_relation or "猫" in worries:
        talk.append("「猫の暮らしも含めて考えると、住まいの正解は少し変わることがあります。」")

    # やらないこと
    avoid = [
        "売却・賃貸・施設入所などの結論をこちらから急がせない。",
        "家族間の温度差を対立として扱わない。",
        "法的判断・医療判断・不動産判断を断定しない。",
        "AI出力をそのまま相談者に渡さず、管理者が確認してから使う。",
    ]

    history_text = "未登録" if history_df.empty else "\n".join([
        f"- {r.get('record_date','')}｜{r.get('record_type','')}｜{r.get('before_status','')}→{r.get('after_status','')}｜{str(r.get('record',''))[:80]}"
        for _, r in history_df.head(5).iterrows()
    ])

    advice = f"""【管理者向け：にゃんとも対応アドバイス案】

作成日：{today_text()}
対象案件：{c.get('case_title','')}
相談者：{c.get('name','')}（{c.get('area','')}）
現在ステータス：{status}
次回確認日：{next_date or '未設定'}

━━━━━━━━━━━━━━━━━━━━
1. 現時点の見立て
━━━━━━━━━━━━━━━━━━━━
この案件は、現時点では「{status}」の段階です。
相談者の状態は「{c.get('current_state','')}」、住まいは「{house_state}」、猫との関係は「{cat_relation}」として記録されています。

この段階では、結論を出すことよりも、
「何が事実で、何が未確定で、何を保留してよいか」を分けることが重要です。

━━━━━━━━━━━━━━━━━━━━
2. 管理者が注意して見る点
━━━━━━━━━━━━━━━━━━━━
{chr(10).join([f"・{x}" for x in attention])}

━━━━━━━━━━━━━━━━━━━━
3. まだ確認した方がよい不足情報
━━━━━━━━━━━━━━━━━━━━
{chr(10).join([f"・{x}" for x in missing]) if missing else "・大きな不足項目は少ないです。ただし、相談者の気持ちの変化は次回も確認してください。"}

━━━━━━━━━━━━━━━━━━━━
4. 次回対応の具体案
━━━━━━━━━━━━━━━━━━━━
{chr(10).join([f"・{x}" for x in next_actions])}

━━━━━━━━━━━━━━━━━━━━
5. 相談者への声かけ案
━━━━━━━━━━━━━━━━━━━━
{chr(10).join([f"・{x}" for x in talk])}

━━━━━━━━━━━━━━━━━━━━
6. 今回あえてやらないこと
━━━━━━━━━━━━━━━━━━━━
{chr(10).join([f"・{x}" for x in avoid])}

━━━━━━━━━━━━━━━━━━━━
7. 最近の履歴
━━━━━━━━━━━━━━━━━━━━
{history_text}

━━━━━━━━━━━━━━━━━━━━
8. 管理者メモ
━━━━━━━━━━━━━━━━━━━━
この出力は、相談者への対応を決めるための補助メモです。
最終判断は管理者が行い、相談者に結論を急がせない前提で使用してください。
"""
    return advice


def build_management_advice_prompt(case_id):
    memo = build_case_memo(case_id)
    return f"""あなたは、にゃんとも相談管理システムの管理者補助AIです。
以下の案件情報をもとに、管理者が相談者へどう対応するかの「対応アドバイス案」を作成してください。

【にゃんとも事業の前提】
・解決を急がせない。
・売らない、急がせない、決めさせない。
・相談者の判断を奪わない。
・保留と放置を分ける。
・人、住まい、猫、家族の温度差を分断せずに扱う。
・結論ではなく、次に確認することを整理する。
・法的判断、医療判断、不動産判断を断定しない。
・家族間の対立を煽らない。
・相談者にそのまま渡す文章ではなく、管理者向けの内部助言として書く。

【出力形式】
1. 現時点の見立て
2. 管理者が注意して見る点
3. まだ確認した方がよい不足情報
4. 次回対応の具体案
5. 相談者への声かけ案
6. 今回あえてやらないこと
7. 管理者メモ

【案件情報】
{memo}
"""


def page_ai_pdf():
    st.subheader("🤖 AI要約・管理者向け対応アドバイス・PDF出力")
    case_id = select_case_widget("ai_pdf_case_select", include_closed=True)
    if not case_id:
        return

    memo = build_case_memo(case_id)
    advice = build_management_advice(case_id)
    prompt = build_management_advice_prompt(case_id)

    ai_tabs = st.tabs(["管理者向け対応アドバイス", "AI用プロンプト", "案件統合メモ", "AI要約履歴", "PDF"])

    with ai_tabs[0]:
        st.caption("案件ごとの現時点情報をもとに、にゃんとも事業の方針に沿って管理者向けの対応案を自動整理します。")
        st.warning("この出力は判断代行ではありません。相談者へそのまま渡さず、管理者が確認してから使用してください。")
        st.text_area("管理者向け対応アドバイス案", advice, height=620)

        with st.form(f"save_auto_advice_{case_id}"):
            note = st.text_area("保存時メモ", placeholder="例：次回面談前に確認、家族温度差に注意 など")
            submitted = st.form_submit_button("この対応アドバイス案をAI履歴に保存", disabled=not has_perm("ai"))
            if submitted:
                c = get_case_full(case_id)
                execute("""
                    INSERT INTO ai_summaries(summary_id, case_id, client_id, created_at, created_by,
                                             summary_type, source_memo, ai_prompt, ai_result, note)
                    VALUES(:summary_id, :case_id, :client_id, :created_at, :created_by,
                           :summary_type, :source_memo, :ai_prompt, :ai_result, :note)
                """, {
                    "summary_id": make_id("ai"),
                    "case_id": case_id,
                    "client_id": c["client_id"],
                    "created_at": now_text(),
                    "created_by": (current_user() or {}).get("username", ""),
                    "summary_type": "管理者向け対応アドバイス",
                    "source_memo": memo,
                    "ai_prompt": prompt,
                    "ai_result": advice,
                    "note": note,
                })
                add_audit_log("save_management_advice", "case", case_id, "管理者向け対応アドバイス保存")
                st.success("対応アドバイス案をAI履歴に保存しました。")
                st.rerun()

    with ai_tabs[1]:
        st.caption("外部AIに貼り付ける場合のプロンプトです。個人情報は必要に応じて伏せてください。")
        st.text_area("管理者向け対応アドバイス生成プロンプト", prompt, height=620)

        st.markdown("### 外部AIの回答を保存")
        with st.form(f"save_external_ai_advice_{case_id}"):
            summary_type = st.selectbox("保存種別", ["外部AI対応アドバイス", "内部整理", "相談者向け要約", "家族共有用", "次回確認用", "その他"])
            ai_result = st.text_area("外部AIで作成した文章を貼り付け", height=260)
            note = st.text_area("補足メモ")
            submitted = st.form_submit_button("外部AI結果を履歴保存", disabled=not has_perm("ai"))
            if submitted:
                c = get_case_full(case_id)
                execute("""
                    INSERT INTO ai_summaries(summary_id, case_id, client_id, created_at, created_by,
                                             summary_type, source_memo, ai_prompt, ai_result, note)
                    VALUES(:summary_id, :case_id, :client_id, :created_at, :created_by,
                           :summary_type, :source_memo, :ai_prompt, :ai_result, :note)
                """, {
                    "summary_id": make_id("ai"),
                    "case_id": case_id,
                    "client_id": c["client_id"],
                    "created_at": now_text(),
                    "created_by": (current_user() or {}).get("username", ""),
                    "summary_type": summary_type,
                    "source_memo": memo,
                    "ai_prompt": prompt,
                    "ai_result": ai_result,
                    "note": note,
                })
                add_audit_log("save_external_ai_advice", "case", case_id, summary_type)
                st.success("AI結果を履歴保存しました。")
                st.rerun()

    with ai_tabs[2]:
        st.text_area("案件統合メモ", memo, height=620)

    with ai_tabs[3]:
        hist = fetch_df("""
            SELECT created_at, created_by, summary_type, ai_result, note, summary_id
            FROM ai_summaries
            WHERE case_id=:case_id
            ORDER BY created_at DESC
        """, {"case_id": case_id})
        st.dataframe(hist, use_container_width=True)
        if not hist.empty:
            selected = st.selectbox(
                "内容を確認する履歴",
                [f"{r['created_at']}｜{r['summary_type']}｜{r['summary_id']}" for _, r in hist.iterrows()],
                key=f"ai_hist_select_{case_id}"
            )
            summary_id = selected_id_from_label(selected)
            row = fetch_one("SELECT * FROM ai_summaries WHERE summary_id=:summary_id", {"summary_id": summary_id})
            if row:
                st.text_area("保存済みAI結果", row.get("ai_result", ""), height=420)
                st.text_area("保存時メモ", row.get("note", ""), height=120)

    with ai_tabs[4]:
        st.download_button("案件統合メモPDFをダウンロード", make_pdf_bytes(memo), file_name=f"nyantomo_memo_{case_id}.pdf", mime="application/pdf")
        st.download_button("管理者向け対応アドバイスPDFをダウンロード", make_pdf_bytes(advice), file_name=f"nyantomo_advice_{case_id}.pdf", mime="application/pdf")


def page_search_update_delete():
    st.subheader("🔎 検索・更新・削除")
    st.caption("Ver2.5.1では、SQLiteの各テーブルを検索し、主要項目を画面から更新できます。")

    table_map = {
        "相談者": "clients",
        "案件": "cases",
        "相談履歴": "history",
        "空き家": "properties",
        "猫": "cats",
        "家族": "family",
        "写真": "photos",
        "AI履歴": "ai_summaries",
        "LINE履歴": "line_messages",
    }

    id_cols = {
        "clients": "client_id",
        "cases": "case_id",
        "history": "history_id",
        "properties": "property_id",
        "cats": "cat_id",
        "family": "family_id",
        "photos": "photo_id",
        "ai_summaries": "summary_id",
        "line_messages": "message_id",
    }

    table_label = st.selectbox("対象データ", list(table_map.keys()), key="sud_table_select")
    table = table_map[table_label]
    id_col = id_cols[table]

    df = fetch_df(f"SELECT * FROM {table}")
    keyword = st.text_input("検索キーワード", placeholder="名前・住所・猫の名前・メモなど", key=f"sud_keyword_{table}")
    if keyword and not df.empty:
        df = df[df.astype(str).apply(lambda row: row.str.contains(keyword, case=False, na=False).any(), axis=1)]

    st.write(f"検索結果：{len(df)}件")
    st.dataframe(df, use_container_width=True)

    if df.empty:
        st.info("更新・削除できるデータがありません。")
        return

    st.divider()
    st.markdown("### 更新")
    if not has_perm("write"):
        st.warning("閲覧権限のみのため、更新はできません。")
    else:
        def make_label(row, cols):
            parts = []
            for col in cols:
                v = str(row.get(col, ""))
                if v:
                    parts.append(v[:30])
            parts.append(str(row.get(id_col, "")))
            return "｜".join(parts)

        label_cols_map = {
            "clients": ["name", "area"],
            "cases": ["case_title", "status"],
            "history": ["record_date", "record_type"],
            "properties": ["property_name", "address"],
            "cats": ["cat_name", "current_life"],
            "family": ["person_name", "relation"],
            "photos": ["photo_type", "original_filename"],
            "ai_summaries": ["created_at", "summary_type"],
            "line_messages": ["created_at", "send_status"],
        }

        labels = [make_label(r, label_cols_map.get(table, [])) for _, r in df.iterrows()]
        selected = st.selectbox("更新するデータ", labels, key=f"sud_update_select_{table}")
        record_id = selected_id_from_label(selected)

        row = fetch_one(f"SELECT * FROM {table} WHERE {id_col}=:id", {"id": record_id})
        if not row:
            st.error("対象データが見つかりません。")
        else:
            with st.form(f"sud_update_form_{table}_{record_id}"):
                update_values = {}

                if table == "clients":
                    st.text_input("client_id", value=row.get("client_id", ""), disabled=True)
                    st.text_input("登録日時", value=row.get("created_at", ""), disabled=True)
                    c1, c2 = st.columns(2)
                    with c1:
                        update_values["name"] = st.text_input("お名前", value=row.get("name", ""))
                        update_values["age_group"] = st.selectbox("年代", AGE_OPTIONS, index=option_index(AGE_OPTIONS, row.get("age_group", "")))
                        update_values["area"] = st.text_input("地域", value=row.get("area", ""))
                    with c2:
                        update_values["contact_method"] = st.selectbox("連絡方法", CONTACT_OPTIONS, index=option_index(CONTACT_OPTIONS, row.get("contact_method", "")))
                        update_values["position"] = st.selectbox("相談者の立場", POSITION_OPTIONS, index=option_index(POSITION_OPTIONS, row.get("position", "")))
                        update_values["note"] = st.text_area("備考", value=row.get("note", ""))

                elif table == "cases":
                    st.text_input("case_id", value=row.get("case_id", ""), disabled=True)
                    st.text_input("client_id", value=row.get("client_id", ""), disabled=True)
                    c1, c2 = st.columns(2)
                    with c1:
                        update_values["case_title"] = st.text_input("案件名", value=row.get("case_title", ""))
                        update_values["case_type"] = st.selectbox("案件種別", CASE_TYPES, index=option_index(CASE_TYPES, row.get("case_type", "")))
                        update_values["status"] = st.selectbox("現在ステータス", STATUS_ORDER, index=option_index(STATUS_ORDER, row.get("status", "")))
                        update_values["consult_date"] = date_or_blank(st.date_input("相談日", value=parse_date_safe(row.get("consult_date")), key=f"sud_consult_{record_id}"))
                        update_values["next_check_date"] = date_or_blank(st.date_input("次回確認日", value=parse_date_safe(row.get("next_check_date")), key=f"sud_next_{record_id}"))
                    with c2:
                        update_values["current_state"] = st.selectbox("今いちばん近い状態", CURRENT_STATE_OPTIONS, index=option_index(CURRENT_STATE_OPTIONS, row.get("current_state", "")))
                        update_values["house_state"] = st.selectbox("住まいの状態", HOUSE_STATE_OPTIONS, index=option_index(HOUSE_STATE_OPTIONS, row.get("house_state", "")))
                        update_values["cat_relation"] = st.selectbox("猫との関係", CAT_RELATION_OPTIONS, index=option_index(CAT_RELATION_OPTIONS, row.get("cat_relation", "")))
                        update_values["family_gap"] = st.selectbox("家族との温度差", FAMILY_GAP_OPTIONS, index=option_index(FAMILY_GAP_OPTIONS, row.get("family_gap", "")))
                        update_values["pressure"] = st.selectbox("急がされている感じ", PRESSURE_OPTIONS, index=option_index(PRESSURE_OPTIONS, row.get("pressure", "")))
                    update_values["worries"] = list_to_text(st.multiselect("気になること", WORRY_OPTIONS, default=[x for x in text_to_list(row.get("worries", "")) if x in WORRY_OPTIONS]))
                    update_values["not_decide"] = st.text_area("今は決めたくないこと", value=row.get("not_decide", ""))
                    update_values["first_check"] = st.text_area("まず確認したいこと", value=row.get("first_check", ""))
                    update_values["free_memo"] = st.text_area("自由メモ", value=row.get("free_memo", ""))
                    update_values["internal_memo"] = st.text_area("内部メモ", value=row.get("internal_memo", ""))
                    update_values["next_check"] = st.text_area("次回確認すること", value=row.get("next_check", ""))
                    update_values["closed_date"] = st.text_input("終了日", value=row.get("closed_date", ""))
                    update_values["close_reason"] = st.text_input("終了理由", value=row.get("close_reason", ""))
                    update_values["final_memo"] = st.text_area("最終メモ", value=row.get("final_memo", ""))
                    update_values["reopen_possibility"] = st.selectbox("再相談可能性", REOPEN_OPTIONS, index=option_index(REOPEN_OPTIONS, row.get("reopen_possibility", "")))

                elif table == "history":
                    st.text_input("history_id", value=row.get("history_id", ""), disabled=True)
                    st.text_input("case_id", value=row.get("case_id", ""), disabled=True)
                    st.text_input("client_id", value=row.get("client_id", ""), disabled=True)
                    c1, c2 = st.columns(2)
                    with c1:
                        update_values["record_date"] = date_or_blank(st.date_input("記録日", value=parse_date_safe(row.get("record_date")), key=f"sud_hist_date_{record_id}"))
                        update_values["record_type"] = st.selectbox("記録種別", ["相談", "電話", "LINE", "メール", "面談", "現地確認", "状態変更", "終了確認", "その他"], index=option_index(["相談", "電話", "LINE", "メール", "面談", "現地確認", "状態変更", "終了確認", "その他"], row.get("record_type", "")))
                    with c2:
                        update_values["before_status"] = st.selectbox("状態変更前", [""] + STATUS_ORDER, index=option_index([""] + STATUS_ORDER, row.get("before_status", "")))
                        update_values["after_status"] = st.selectbox("状態変更後", [""] + STATUS_ORDER, index=option_index([""] + STATUS_ORDER, row.get("after_status", "")))
                    update_values["record"] = st.text_area("相談記録", value=row.get("record", ""))
                    update_values["next_action"] = st.text_area("次回アクション", value=row.get("next_action", ""))
                    update_values["internal_memo"] = st.text_area("内部メモ", value=row.get("internal_memo", ""))

                elif table == "properties":
                    st.text_input("property_id", value=row.get("property_id", ""), disabled=True)
                    st.text_input("case_id", value=row.get("case_id", ""), disabled=True)
                    c1, c2 = st.columns(2)
                    with c1:
                        update_values["property_name"] = st.text_input("物件名", value=row.get("property_name", ""))
                        update_values["address"] = st.text_input("所在地", value=row.get("address", ""))
                        update_values["property_status"] = st.selectbox("物件状態", PROPERTY_STATUS_OPTIONS, index=option_index(PROPERTY_STATUS_OPTIONS, row.get("property_status", "")))
                        update_values["vacant_status"] = st.selectbox("空き家状態", VACANT_STATUS_OPTIONS, index=option_index(VACANT_STATUS_OPTIONS, row.get("vacant_status", "")))
                    with c2:
                        update_values["key_hold"] = st.selectbox("鍵預かり", KEY_HOLD_OPTIONS, index=option_index(KEY_HOLD_OPTIONS, row.get("key_hold", "")))
                        update_values["neighborhood_anxiety"] = st.selectbox("近隣不安", NEIGHBORHOOD_OPTIONS, index=option_index(NEIGHBORHOOD_OPTIONS, row.get("neighborhood_anxiety", "")))
                        update_values["management_frequency"] = st.selectbox("管理頻度", FREQUENCY_OPTIONS, index=option_index(FREQUENCY_OPTIONS, row.get("management_frequency", "")))
                    update_values["memo"] = st.text_area("メモ", value=row.get("memo", ""))

                elif table == "cats":
                    st.text_input("cat_id", value=row.get("cat_id", ""), disabled=True)
                    st.text_input("case_id", value=row.get("case_id", ""), disabled=True)
                    c1, c2 = st.columns(2)
                    with c1:
                        update_values["cat_name"] = st.text_input("猫の名前", value=row.get("cat_name", ""))
                        update_values["age"] = st.text_input("年齢", value=row.get("age", ""))
                        update_values["count"] = st.text_input("頭数", value=row.get("count", ""))
                        update_values["current_life"] = st.selectbox("現在の暮らし", CAT_LIFE_OPTIONS, index=option_index(CAT_LIFE_OPTIONS, row.get("current_life", "")))
                    with c2:
                        update_values["concerns"] = st.text_area("気になること", value=row.get("concerns", ""))
                        update_values["place_candidate"] = st.text_area("預け先候補", value=row.get("place_candidate", ""))
                        update_values["memo"] = st.text_area("メモ", value=row.get("memo", ""))

                elif table == "family":
                    st.text_input("family_id", value=row.get("family_id", ""), disabled=True)
                    st.text_input("case_id", value=row.get("case_id", ""), disabled=True)
                    c1, c2 = st.columns(2)
                    with c1:
                        update_values["person_name"] = st.text_input("関係者名", value=row.get("person_name", ""))
                        update_values["relation"] = st.text_input("続柄", value=row.get("relation", ""))
                        update_values["contact_ok"] = st.selectbox("連絡可否", CONTACT_OK_OPTIONS, index=option_index(CONTACT_OK_OPTIONS, row.get("contact_ok", "")))
                    with c2:
                        update_values["temperature"] = st.selectbox("温度感", TEMP_OPTIONS, index=option_index(TEMP_OPTIONS, row.get("temperature", "")))
                        update_values["relation_memo"] = st.text_area("関係メモ", value=row.get("relation_memo", ""))

                elif table == "photos":
                    st.text_input("photo_id", value=row.get("photo_id", ""), disabled=True)
                    st.text_input("case_id", value=row.get("case_id", ""), disabled=True)
                    update_values["photo_type"] = st.selectbox("写真種別", PHOTO_TYPE_OPTIONS, index=option_index(PHOTO_TYPE_OPTIONS, row.get("photo_type", "")))
                    update_values["original_filename"] = st.text_input("元ファイル名", value=row.get("original_filename", ""))
                    st.text_input("保存先", value=row.get("saved_path", ""), disabled=True)
                    update_values["description"] = st.text_area("説明", value=row.get("description", ""))

                elif table == "ai_summaries":
                    st.text_input("summary_id", value=row.get("summary_id", ""), disabled=True)
                    st.text_input("case_id", value=row.get("case_id", ""), disabled=True)
                    update_values["summary_type"] = st.text_input("要約種別", value=row.get("summary_type", ""))
                    update_values["ai_result"] = st.text_area("AI結果", value=row.get("ai_result", ""), height=260)
                    update_values["note"] = st.text_area("補足メモ", value=row.get("note", ""))

                elif table == "line_messages":
                    st.text_input("message_id", value=row.get("message_id", ""), disabled=True)
                    st.text_input("case_id", value=row.get("case_id", ""), disabled=True)
                    update_values["to_target"] = st.text_input("送信先ID", value=row.get("to_target", ""))
                    update_values["message_text"] = st.text_area("送信文", value=row.get("message_text", ""), height=220)
                    update_values["send_status"] = st.selectbox("送信状態", ["下書き保存", "送信予定", "送信済", "送信失敗", "中止"], index=option_index(["下書き保存", "送信予定", "送信済", "送信失敗", "中止"], row.get("send_status", "")))
                    update_values["response_memo"] = st.text_area("反応・メモ", value=row.get("response_memo", ""))

                submitted = st.form_submit_button("この内容で更新する")
                if submitted:
                    update_values = {k: ("" if v is None else v) for k, v in update_values.items()}
                    if table == "cases":
                        update_values["updated_at"] = now_text()
                    set_sql = ", ".join([f"{k}=:{k}" for k in update_values.keys()])
                    params = dict(update_values)
                    params["id"] = record_id
                    execute(f"UPDATE {table} SET {set_sql} WHERE {id_col}=:id", params)

                    # 関連データ更新時は親案件の updated_at も更新
                    if table in ["history", "properties", "cats", "family", "photos", "ai_summaries", "line_messages"]:
                        related_case_id = row.get("case_id", "")
                        if related_case_id:
                            execute("UPDATE cases SET updated_at=:updated_at WHERE case_id=:case_id", {"updated_at": now_text(), "case_id": related_case_id})

                    add_audit_log("update_record", table, record_id, "検索・更新・削除画面から更新")
                    st.success("更新しました。")
                    st.rerun()

    st.divider()
    st.markdown("### 削除")
    st.warning("削除は元に戻せません。削除前に必ずバックアップしてください。")

    if not has_perm("delete"):
        st.info("削除は管理者権限のみ可能です。")
        return

    delete_id = st.selectbox("削除するID", df[id_col].astype(str).tolist(), key=f"delete_{table}")
    confirm = st.checkbox("本当に削除します", key=f"confirm_delete_{table}")
    if st.button("削除する", type="primary", disabled=not confirm):
        removed_files = 0
        if table == "cases":
            removed_files = delete_photo_files_for_case(delete_id)
        elif table == "photos":
            removed_files = delete_photo_file_by_id(delete_id)
        execute(f"DELETE FROM {table} WHERE {id_col}=:id", {"id": delete_id})
        add_audit_log("delete_record", table, delete_id, f"removed_photo_files={removed_files}")
        st.success(f"削除しました。写真ファイル削除：{removed_files}件")
        st.rerun()


def page_data_management():
    st.subheader("📦 データ管理")
    st.caption("SQLite版のバックアップ・復元・Excel移行を行います。")

    c1, c2, c3 = st.columns(3)
    c1.metric("相談者", table_count("clients"))
    c2.metric("案件", table_count("cases"))
    c3.metric("履歴", table_count("history"))

    st.markdown("### リレーション整合性チェック")
    health_df = relation_health_check()
    if health_df.empty:
        st.success("リレーションに大きな問題は見つかりません。")
    else:
        st.warning(f"確認が必要なデータが {len(health_df)} 件あります。")
        st.dataframe(health_df, use_container_width=True)
        if st.button("case_idを基準にclient_id不一致を自動修復する"):
            result = repair_child_client_ids()
            st.success(f"修復しました：{result}")
            st.rerun()

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
        restored = safe_restore_backup_zip(restore_file)
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




def page_user_management():
    st.subheader("🔐 ログイン・権限管理")
    if not require_admin():
        return

    st.markdown("### ユーザー追加")
    with st.form("add_user_form"):
        username = st.text_input("ユーザー名")
        display_name = st.text_input("表示名")
        password = st.text_input("初期パスワード", type="password")
        role = st.selectbox("権限", ROLE_OPTIONS)
        submitted = st.form_submit_button("ユーザーを追加")
        if submitted:
            if not username or not password:
                st.error("ユーザー名とパスワードを入力してください。")
            else:
                try:
                    execute("""
                        INSERT INTO users(user_id, username, display_name, password_hash, role, is_active, created_at)
                        VALUES(:user_id, :username, :display_name, :password_hash, :role, '1', :created_at)
                    """, {
                        "user_id": make_id("user"),
                        "username": username,
                        "display_name": display_name,
                        "password_hash": hash_password(password),
                        "role": role,
                        "created_at": now_text(),
                    })
                    add_audit_log("create_user", "user", username, role)
                    st.success("ユーザーを追加しました。")
                    st.rerun()
                except Exception as e:
                    st.error(f"追加できませんでした：{e}")

    st.markdown("### ユーザー一覧")
    users = fetch_df("SELECT user_id, username, display_name, role, is_active, created_at FROM users ORDER BY created_at DESC")
    st.dataframe(users, use_container_width=True)

    st.markdown("### 権限変更・停止")
    if not users.empty:
        selected = st.selectbox("対象ユーザー", [f"{r['username']}｜{r['role']}｜{r['user_id']}" for _, r in users.iterrows()])
        user_id = selected_id_from_label(selected)
        u = fetch_one("SELECT * FROM users WHERE user_id=:user_id", {"user_id": user_id})
        with st.form(f"edit_user_{user_id}"):
            new_role = st.selectbox("権限", ROLE_OPTIONS, index=option_index(ROLE_OPTIONS, u.get("role", "閲覧者")))
            active = st.selectbox("有効状態", ["1", "0"], index=0 if u.get("is_active", "1") == "1" else 1, format_func=lambda x: "有効" if x == "1" else "停止")
            new_password = st.text_input("新パスワード（変更時のみ）", type="password")
            submitted = st.form_submit_button("更新")
            if submitted:
                if new_password:
                    execute("UPDATE users SET role=:role, is_active=:active, password_hash=:ph WHERE user_id=:user_id", {
                        "role": new_role, "active": active, "ph": hash_password(new_password), "user_id": user_id
                    })
                else:
                    execute("UPDATE users SET role=:role, is_active=:active WHERE user_id=:user_id", {
                        "role": new_role, "active": active, "user_id": user_id
                    })
                add_audit_log("update_user", "user", user_id, f"{new_role}/{active}")
                st.success("更新しました。")
                st.rerun()

    st.markdown("### 操作ログ")
    logs = fetch_df("SELECT created_at, username, action, target_type, target_id, detail FROM audit_logs ORDER BY created_at DESC LIMIT 300")
    st.dataframe(logs, use_container_width=True)


def page_line_linkage():
    st.subheader("📱 LINE連携")
    st.caption("Ver2.1では、LINE送信用の設定・送信文作成・送信履歴保存まで対応します。実送信はチャネルアクセストークン設定後の運用を想定します。")

    if not has_perm("line"):
        st.warning("LINE連携は管理者または編集者のみ使用できます。")
        return

    st.markdown("### LINE設定")
    setting = fetch_one("SELECT * FROM line_settings LIMIT 1")
    with st.form("line_settings_form"):
        token = st.text_input("チャネルアクセストークン", value=(setting or {}).get("channel_access_token", ""), type="password")
        secret = st.text_input("チャネルシークレット", value=(setting or {}).get("channel_secret", ""), type="password")
        default_to = st.text_input("送信先ID（userId/groupIdなど）", value=(setting or {}).get("default_to", ""))
        enabled = st.selectbox("連携状態", ["0", "1"], index=1 if (setting or {}).get("enabled") == "1" else 0, format_func=lambda x: "有効" if x == "1" else "無効")
        note = st.text_area("設定メモ", value=(setting or {}).get("note", ""))
        submitted = st.form_submit_button("LINE設定を保存")
        if submitted:
            execute("DELETE FROM line_settings", {})
            execute("""
                INSERT INTO line_settings(setting_id, channel_access_token, channel_secret, default_to, enabled, note, updated_at)
                VALUES(:setting_id, :token, :secret, :default_to, :enabled, :note, :updated_at)
            """, {
                "setting_id": "line_main",
                "token": token,
                "secret": secret,
                "default_to": default_to,
                "enabled": enabled,
                "note": note,
                "updated_at": now_text(),
            })
            add_audit_log("update_line_settings", "line", "line_main", "LINE設定更新")
            st.success("LINE設定を保存しました。")
            st.rerun()

    st.markdown("### 案件別LINE文作成・履歴保存")
    case_id = select_case_widget("line_case_select", include_closed=True)
    if case_id:
        c = get_case_full(case_id)
        default_message = f"""【にゃんとも相談室】
{c.get('name','')}様

前回のご相談内容について、次回確認予定日が近づいています。
次回確認予定日：{c.get('next_check_date','未設定')}

確認したいこと：
{c.get('next_check','')}

※この連絡は、判断を急がせるものではなく、状況確認のためのものです。
"""
        with st.form(f"line_message_form_{case_id}"):
            to_target = st.text_input("送信先ID", value=(setting or {}).get("default_to", "") if setting else "")
            message_text = st.text_area("送信文", value=default_message, height=260)
            save_only = st.form_submit_button("送信履歴として保存")
            if save_only:
                execute("""
                    INSERT INTO line_messages(message_id, case_id, client_id, created_at, created_by,
                                              to_target, message_text, send_status, response_memo)
                    VALUES(:message_id, :case_id, :client_id, :created_at, :created_by,
                           :to_target, :message_text, :send_status, :response_memo)
                """, {
                    "message_id": make_id("line"),
                    "case_id": case_id,
                    "client_id": c["client_id"],
                    "created_at": now_text(),
                    "created_by": (current_user() or {}).get("username", ""),
                    "to_target": to_target,
                    "message_text": message_text,
                    "send_status": "下書き保存",
                    "response_memo": "",
                })
                add_audit_log("save_line_draft", "case", case_id, "LINE下書き保存")
                st.success("LINE送信文を履歴保存しました。")
                st.rerun()

    st.markdown("### LINE送信履歴")
    df = fetch_df("""
        SELECT lm.created_at, cl.name AS 相談者, c.case_title AS 案件名, lm.to_target, lm.send_status, lm.message_text, lm.response_memo, lm.message_id
        FROM line_messages lm
        LEFT JOIN cases c ON lm.case_id = c.case_id
        LEFT JOIN clients cl ON lm.client_id = cl.client_id
        ORDER BY lm.created_at DESC
    """)
    st.dataframe(df, use_container_width=True)


def page_ai_history():
    st.subheader("🧠 AI要約履歴")
    df = fetch_df("""
        SELECT a.created_at, cl.name AS 相談者, c.case_title AS 案件名, a.created_by, a.summary_type, a.ai_result, a.note, a.summary_id, a.case_id
        FROM ai_summaries a
        JOIN cases c ON a.case_id = c.case_id
        JOIN clients cl ON a.client_id = cl.client_id
        ORDER BY a.created_at DESC
    """)
    keyword = st.text_input("AI履歴検索", placeholder="相談者・案件・要約内容など")
    if keyword and not df.empty:
        df = df[df.astype(str).apply(lambda row: row.str.contains(keyword, case=False, na=False).any(), axis=1)]
    st.dataframe(df, use_container_width=True)



def page_relation_maintenance():
    st.subheader("🧰 リレーション整合性")
    st.caption("case_id・client_idの不一致、孤立データ、写真ファイル欠損を確認します。")

    ensure_indexes()

    check_df = relation_check_df()
    if check_df.empty:
        st.success("リレーションの大きな問題は見つかりません。")
    else:
        st.warning(f"確認が必要な項目が {len(check_df)} 件あります。")
        st.dataframe(check_df, use_container_width=True)

    st.markdown("### 自動修復")
    st.caption("client_id不一致のみ、親案件のclient_idに揃えて自動修復できます。孤立データや写真欠損は内容確認が必要です。")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("client_id不一致を自動修復", disabled=not has_perm("admin")):
            fixed = repair_client_id_mismatches()
            st.success(f"{fixed}件を修復しました。")
            st.rerun()

    with c2:
        st.warning("孤立データ削除は復元ミス時のみ使用してください。")
        confirm = st.checkbox("孤立データ削除を許可する", key="confirm_delete_orphans")
        if st.button("親案件のない孤立データを削除", disabled=(not confirm or not has_perm("admin"))):
            deleted = delete_orphan_related_records()
            st.success(f"{deleted}件の孤立データを削除しました。")
            st.rerun()

    st.markdown("### 件数確認")
    counts = pd.DataFrame([
        {"テーブル": "clients", "件数": table_count("clients")},
        {"テーブル": "cases", "件数": table_count("cases")},
        {"テーブル": "history", "件数": table_count("history")},
        {"テーブル": "properties", "件数": table_count("properties")},
        {"テーブル": "cats", "件数": table_count("cats")},
        {"テーブル": "family", "件数": table_count("family")},
        {"テーブル": "photos", "件数": table_count("photos")},
        {"テーブル": "ai_summaries", "件数": table_count("ai_summaries")},
        {"テーブル": "line_messages", "件数": table_count("line_messages")},
    ])
    st.dataframe(counts, use_container_width=True)


# -----------------------------
# 起動
# -----------------------------
init_db()
ensure_indexes()
ensure_default_admin()

if not current_user():
    login_screen()
    st.stop()

render_top_nav()
logout_button()

st.title("🐾 にゃんとも相談管理システム Ver2.5.1.1（伴走支援エラー修正版）")
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
    "🤝 伴走支援",
    "📚 相談履歴",
    "⏸ 保留案件一覧",
    "🏠 空き家カード",
    "🗺 GoogleMap",
    "🐈 猫情報カード",
    "👪 家族関係メモ",
    "📷 写真管理",
    "🤖 AI助言/PDF",
    "🔎 検索・更新・削除",
    "📦 データ管理",
    "🧰 リレーション整合性",
    "🔐 権限管理",
    "📱 LINE連携",
    "🧠 AI履歴",
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
    page_companion_support()
with tabs[5]:
    page_history()
with tabs[6]:
    page_hold_list()
with tabs[7]:
    page_property()
with tabs[8]:
    page_map()
with tabs[9]:
    page_cat()
with tabs[10]:
    page_family()
with tabs[11]:
    page_photos()
with tabs[12]:
    page_ai_pdf()
with tabs[13]:
    page_search_update_delete()
with tabs[14]:
    page_data_management()
with tabs[15]:
    page_relation_maintenance()
with tabs[16]:
    page_user_management()
with tabs[17]:
    page_line_linkage()
with tabs[18]:
    page_ai_history()
