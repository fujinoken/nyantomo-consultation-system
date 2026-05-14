import streamlit as st
import pandas as pd
from datetime import date, datetime
from pathlib import Path
import uuid

# =========================================================
# にゃんとも相談管理システム Ver1.1 基礎設計版
# ---------------------------------------------------------
# 目的：
# ・今は Excel 保存で安定運用
# ・将来 SQLite 化しやすいように「ID」「カード」「状態遷移」を先に入れる
# ・診断しない／決めさせない／保留を支える記録システム
# =========================================================

st.set_page_config(
    page_title="にゃんとも相談管理システム",
    page_icon="🐾",
    layout="wide"
)

DATA_FILE = Path("nyantomo_consultation_data.xlsx")

# ---------------------------------------------------------
# シート定義
# ---------------------------------------------------------
CLIENT_COLUMNS = [
    "client_id",
    "登録日時",
    "お名前",
    "年代",
    "地域",
    "連絡方法",
    "相談者の立場",
    "備考",
]

CASE_COLUMNS = [
    "case_id",
    "client_id",
    "登録日時",
    "相談日",
    "案件名",
    "案件種別",
    "現在ステータス",
    "今いちばん近い状態",
    "住まいの状態",
    "猫との関係",
    "家族との温度差",
    "急がされている感じ",
    "気になること",
    "今は決めたくないこと",
    "まず確認したいこと",
    "自由メモ",
    "内部メモ",
    "次回確認すること",
]

HISTORY_COLUMNS = [
    "history_id",
    "case_id",
    "client_id",
    "記録日時",
    "記録日",
    "記録種別",
    "状態変更前",
    "状態変更後",
    "相談記録",
    "次回アクション",
    "内部メモ",
]

PROPERTY_COLUMNS = [
    "property_id",
    "case_id",
    "client_id",
    "登録日時",
    "物件名",
    "所在地",
    "物件状態",
    "空き家状態",
    "鍵預かり",
    "近隣不安",
    "管理頻度",
    "メモ",
]

CAT_COLUMNS = [
    "cat_id",
    "case_id",
    "client_id",
    "登録日時",
    "猫の名前",
    "年齢",
    "頭数",
    "現在の暮らし",
    "気になること",
    "預け先候補",
    "メモ",
]

FAMILY_COLUMNS = [
    "family_id",
    "case_id",
    "client_id",
    "登録日時",
    "関係者名",
    "続柄",
    "連絡可否",
    "温度感",
    "関係メモ",
]

PHOTO_COLUMNS = [
    "photo_id",
    "case_id",
    "client_id",
    "登録日時",
    "写真種別",
    "ファイル名",
    "保存先メモ",
    "説明",
]

STATUS_ORDER = [
    "未対応",
    "初回相談前",
    "初回相談済",
    "情報整理中",
    "保留中",
    "見守り中",
    "継続相談",
    "専門家紹介済",
    "終了",
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
    name = row.get("お名前", "")
    area = row.get("地域", "")
    cid = row.get("client_id", "")
    return f"{name}｜{area}｜{cid}"


def get_case_label(row):
    title = row.get("案件名", "")
    status = row.get("現在ステータス", "")
    case_id = row.get("case_id", "")
    return f"{title}｜{status}｜{case_id}"


def selected_id_from_label(label):
    if not label or "｜" not in label:
        return ""
    return label.split("｜")[-1]


data = load_all()

# ---------------------------------------------------------
# 画面
# ---------------------------------------------------------
st.title("🐾 にゃんとも相談管理システム Ver1.1")
st.caption("相談履歴・状態遷移・保留案件・空き家カード・猫情報カードへ拡張できる基礎設計版")

tabs = st.tabs([
    "🧑 相談者登録",
    "📝 案件登録",
    "📚 相談履歴",
    "⏸ 保留案件一覧",
    "🏠 空き家カード",
    "🐈 猫情報カード",
    "👪 家族関係メモ",
    "🧾 相談メモ出力",
    "📦 データ管理",
])

# ---------------------------------------------------------
# 相談者登録
# ---------------------------------------------------------
with tabs[0]:
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

# ---------------------------------------------------------
# 案件登録
# ---------------------------------------------------------
with tabs[1]:
    st.subheader("案件登録")

    if data["clients"].empty:
        st.info("先に相談者を登録してください。")
    else:
        client_labels = [get_client_label(row) for _, row in data["clients"].iterrows()]

        with st.form("case_form"):
            selected_client_label = st.selectbox("相談者を選択", client_labels)
            client_id = selected_id_from_label(selected_client_label)

            col1, col2 = st.columns(2)

            with col1:
                consult_date = st.date_input("相談日", value=date.today())
                case_title = st.text_input("案件名", value="住まいと猫の相談")
                case_type = st.selectbox("案件種別", ["初回相談", "空き家管理", "猫と住まい", "相続前整理", "高齢期の住まい", "その他"])
                status = st.selectbox("現在ステータス", STATUS_ORDER, index=1)
                current_state = st.selectbox(
                    "今いちばん近い状態",
                    [
                        "未選択",
                        "まだ何も決まっていない",
                        "少し考え始めている",
                        "家族と話し始めた",
                        "急かされている感じがある",
                        "誰にも相談していない",
                        "すでに困りごとが出ている",
                    ],
                )

            with col2:
                house_state = st.selectbox(
                    "住まいの状態",
                    [
                        "未選択",
                        "現在住んでいる",
                        "空き家になっている",
                        "近いうちに空き家になりそう",
                        "相続後そのまま",
                        "売却・賃貸を迷っている",
                        "荷物整理が進んでいない",
                    ],
                )
                cat_relation = st.selectbox(
                    "猫との関係",
                    [
                        "未選択",
                        "猫と暮らしている",
                        "家族の猫がいる",
                        "猫を残して入院・施設入所が心配",
                        "これから猫と暮らしたい",
                        "保護猫に関心がある",
                        "猫はいない",
                    ],
                )
                family_gap = st.selectbox("家族との温度差", ["未選択", "特にない", "少しある", "かなりある", "まだ話せていない"])
                pressure = st.selectbox("急がされている感じ", ["未選択", "ない", "少しある", "強くある", "自分でも焦っている"])

            worries = st.multiselect(
                "気になること",
                [
                    "空き家管理",
                    "相続",
                    "売却",
                    "賃貸",
                    "猫の住まい",
                    "高齢期の暮らし",
                    "家族との意見の違い",
                    "お金",
                    "近所への不安",
                    "何から考えればよいか分からない",
                ],
            )

            not_decide = st.text_area("今は決めたくないこと")
            first_check = st.text_area("まず確認したいこと")
            free_memo = st.text_area("自由メモ")
            internal_memo = st.text_area("内部メモ")
            next_check = st.text_area("次回確認すること")

            submitted = st.form_submit_button("案件を登録")

            if submitted:
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

# ---------------------------------------------------------
# 相談履歴
# ---------------------------------------------------------
with tabs[2]:
    st.subheader("相談履歴・状態変更")

    if data["cases"].empty:
        st.info("先に案件を登録してください。")
    else:
        case_labels = [get_case_label(row) for _, row in data["cases"].iterrows()]
        selected_case_label = st.selectbox("案件を選択", case_labels)
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
                    record_date = st.date_input("記録日", value=date.today())
                    record_type = st.selectbox("記録種別", ["相談", "電話", "LINE", "メール", "面談", "現地確認", "状態変更", "その他"])
                with col2:
                    new_status = st.selectbox("状態変更後", STATUS_ORDER, index=STATUS_ORDER.index(current_status) if current_status in STATUS_ORDER else 0)

                record = st.text_area("相談記録")
                next_action = st.text_area("次回アクション")
                internal = st.text_area("内部メモ")

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
        history_df = data["history"][data["history"]["case_id"] == case_id]
        st.dataframe(history_df, use_container_width=True)

# ---------------------------------------------------------
# 保留案件一覧
# ---------------------------------------------------------
with tabs[3]:
    st.subheader("保留案件一覧")

    if data["cases"].empty:
        st.info("案件がありません。")
    else:
        hold_df = data["cases"][data["cases"]["現在ステータス"].isin(["保留中", "情報整理中", "見守り中"])]
        if hold_df.empty:
            st.success("現在、保留・見守り中の案件はありません。")
        else:
            st.dataframe(hold_df, use_container_width=True)

# ---------------------------------------------------------
# 空き家カード
# ---------------------------------------------------------
with tabs[4]:
    st.subheader("空き家カード")

    if data["cases"].empty:
        st.info("先に案件を登録してください。")
    else:
        case_labels = [get_case_label(row) for _, row in data["cases"].iterrows()]
        selected_case_label = st.selectbox("案件選択（空き家カード）", case_labels)
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

# ---------------------------------------------------------
# 猫情報カード
# ---------------------------------------------------------
with tabs[5]:
    st.subheader("猫情報カード")

    if data["cases"].empty:
        st.info("先に案件を登録してください。")
    else:
        case_labels = [get_case_label(row) for _, row in data["cases"].iterrows()]
        selected_case_label = st.selectbox("案件選択（猫情報カード）", case_labels)
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

# ---------------------------------------------------------
# 家族関係メモ
# ---------------------------------------------------------
with tabs[6]:
    st.subheader("家族関係メモ")

    if data["cases"].empty:
        st.info("先に案件を登録してください。")
    else:
        case_labels = [get_case_label(row) for _, row in data["cases"].iterrows()]
        selected_case_label = st.selectbox("案件選択（家族関係メモ）", case_labels)
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

# ---------------------------------------------------------
# 相談メモ出力
# ---------------------------------------------------------
with tabs[7]:
    st.subheader("相談メモ出力")

    if data["cases"].empty:
        st.info("案件がありません。")
    else:
        case_labels = [get_case_label(row) for _, row in data["cases"].iterrows()]
        selected_case_label = st.selectbox("案件選択（メモ出力）", case_labels)
        case_id = selected_id_from_label(selected_case_label)

        case_row = data["cases"][data["cases"]["case_id"] == case_id].iloc[0]
        client_row = data["clients"][data["clients"]["client_id"] == case_row["client_id"]].iloc[0]

        history_df = data["history"][data["history"]["case_id"] == case_id]
        prop_df = data["properties"][data["properties"]["case_id"] == case_id]
        cat_df = data["cats"][data["cats"]["case_id"] == case_id]
        fam_df = data["family"][data["family"]["case_id"] == case_id]

        history_text = ""
        for _, h in history_df.iterrows():
            history_text += f"- {h['記録日']}｜{h['記録種別']}｜{h['状態変更前']} → {h['状態変更後']}｜{h['相談記録']}\n"

        prop_text = ""
        for _, p in prop_df.iterrows():
            prop_text += f"- {p['物件名']}｜{p['所在地']}｜{p['物件状態']}｜{p['空き家状態']}\n"

        cat_text = ""
        for _, c in cat_df.iterrows():
            cat_text += f"- {c['猫の名前']}｜{c['年齢']}｜{c['頭数']}｜{c['現在の暮らし']}\n"

        fam_text = ""
        for _, f in fam_df.iterrows():
            fam_text += f"- {f['関係者名']}｜{f['続柄']}｜{f['温度感']}｜{f['関係メモ']}\n"

        memo = f"""
【にゃんとも相談整理メモ】

作成日：{date.today().strftime('%Y-%m-%d')}

■ 相談者
お名前：{client_row['お名前']}
地域：{client_row['地域']}
年代：{client_row['年代']}
連絡方法：{client_row['連絡方法']}
相談者の立場：{client_row['相談者の立場']}

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
{prop_text if prop_text else '未登録'}

■ 猫情報カード
{cat_text if cat_text else '未登録'}

■ 家族関係メモ
{fam_text if fam_text else '未登録'}

■ 相談履歴
{history_text if history_text else '未登録'}

■ 内部メモ
{case_row['内部メモ']}

■ 次回確認すること
{case_row['次回確認すること']}

※このメモは、判断を急がせず、状況を整理するための内部記録です。
※法的判断・医療判断・不動産判断を断定するものではありません。
"""

        st.text_area("相談整理メモ", memo, height=700)

# ---------------------------------------------------------
# データ管理
# ---------------------------------------------------------
with tabs[8]:
    st.subheader("データ管理")

    st.write("Excel保存ファイル：", DATA_FILE.name)

    if DATA_FILE.exists():
        with open(DATA_FILE, "rb") as f:
            st.download_button(
                "Excelデータをダウンロード",
                data=f,
                file_name="nyantomo_consultation_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    st.divider()

    st.markdown("### 将来拡張メモ")
    st.write(
        """
        このVer1.1は、次の拡張を想定した構造です。

        - 相談履歴管理：historyシート
        - 状態遷移：casesの現在ステータス＋historyの状態変更前後
        - 保留案件一覧：保留中・情報整理中・見守り中を抽出
        - LINE連携：client_id / case_id をキーに外部連携可能
        - AI要約：相談メモ出力をAIに渡す前提
        - PDF出力：相談整理メモをPDF化する前提
        - 空き家カード：propertiesシート
        - 猫情報カード：catsシート
        - 家族関係図：familyシート
        - GoogleMap連携：所在地カラムを使用
        - 写真保存：photosシートを将来利用
        - SQLite化：各シートをそのままテーブル化可能
        """
    )
