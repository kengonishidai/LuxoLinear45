import io
import datetime

import cv2
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components

from PIL import Image

try:
    import japanize_matplotlib  # noqa: F401
except ImportError:
    pass


# =========================================================
# 基本設定
# =========================================================
st.set_page_config(page_title="超音波プローブ点検アプリ", layout="wide")

PAGE_HOME = "🏠 ホーム / プローブ選択"
PAGE_ABOUT = "ℹ️ アプリについて"

PAGE_OPTIONS = [PAGE_HOME, PAGE_ABOUT]

# 点検するプローブ本数の選択肢
PROBE_COUNT_OPTIONS = [1, 2, 3]

CONVEX_PROBES = ["C5-2", "MC10-3", "EC9-3"]

PROBE_OPTIONS = [
    "",
    "X20L",
    "L18-4",
    "L11-3",
    "HL18-4",
    "WL13-3",
    "L14-4",
    "MC10-3",
    "EC9-3",
    "C5-2",
]

# 超音波診断装置本体の選択肢
DEVICE_OPTIONS = ["UX1/LX", "HS2", "MX1", "GX30/GX20"]

PROBE_DEVICE_MAP = {
    "X20L": ["UX1/LX"],
    "L11-3": ["UX1/LX", "HS2", "MX1"],
    "HL18-4": ["UX1/LX", "HS2", "MX1"],
    "WL13-3": ["UX1/LX", "HS2"],
    "L18-4": ["HS2", "MX1", "GX30/GX20"],
    "L14-4": ["HS2", "MX1"],
    "C5-2": ["UX1/LX", "HS2", "MX1", "GX30/GX20"],
    "EC9-3": ["UX1/LX", "HS2", "GX30/GX20"],
    "MC10-3": ["UX1/LX", "HS2", "MX1"],
}

# 装置ごとのSW Ver選択肢(ここに無い装置はSW Ver選択なし)
SW_VERSION_MAP = {
    "UX1/LX": ["V1.03以降", "V1.02以前"],
    "MX1": ["V1.10以降", "V1.00以前"],
}

# プローブ自己診断機能(CSVアップロード・結果表示・総合判定)を
# 表示するSW Ver
DIAGNOSTIC_SW_VERSIONS = ("V1.03以降", "V1.10以降")

SETTING_LABELS = [
    "TGC",
    "Depth (cm)",
    "Focus (mm)",
    "P (Tx Level)",
    "THI",
    "Frequency",
    "BG",
    "DR",
    "B Filter",
    "コントラスト",
    "グレイマップ",
    "音線密度",
    "iXRET",
    "オブリーク強調",
    "フォーカス段数",
    "エッジ強調",
    "B SI-Filter",
    "Compound",
    "Sound Speed",
    "カラー",
    "時間平均",
]


def preset_row(tgc, depth, focus, thi, frequency, bg, dr):
    """設定表の1プローブ分を作成する。"""
    return [
        tgc,
        depth,
        focus,
        100,
        thi,
        frequency,
        bg,
        dr,
        1,
        50,
        1,
        "H",
        "Off",
        "Off",
        1,
        "Off",
        "Off",
        "Off",
        0,
        "Off",
        "最大",
    ]


DEFAULT_PRESET = {
    "設定項目": SETTING_LABELS,
    "C5-2": preset_row("全て中央", 15, "最浅に設定", "On", "Hres", 40, 70),
    "L18-4": preset_row("全て中央", 4, "最浅に設定", "On", "Hres2", 40, 60),
    "L14-4": preset_row("全て中央", 4, "最浅に設定", "On", "Hres", 30, 60),
    "L11-3": preset_row("全て中央", 4, "最浅に設定", "On", "Hgen", 40, 60),
    "HL18-4": preset_row("全て中央", 4, "最浅に設定", "Off", "Res", 40, 60),
    "EC9-3": preset_row("全て中央", 10, "最浅に設定", "On", "Hres", 40, 60),
    "MC10-3": preset_row("全て中央", 4, "最浅に設定", "On", "Hgen", 40, 60),
    "WL13-3": preset_row("全て中央", 4, "最浅に設定", "On", "Hgen", 35, 60),
    "X20L": preset_row("全て中央", 4, "最浅に設定", "On", "Hres2", 40, 60),
}


# =========================================================
# UX1/LX：Luxo本体条件設定.xlsx
# =========================================================
UX1LX_SETTING_LABELS = SETTING_LABELS + ["2D表示角"]


def ux1lx_preset_row(depth, focus, frequency, bg, dr, display_angle=""):
    row = preset_row(
        "全てセンター", depth, focus, "On", frequency, bg, dr
    )
    row[9] = 50  # コントラスト：全プローブ共通
    row[10] = 1  # グレイマップ：全プローブ共通
    row[20] = 10  # 時間平均
    row.append(display_angle)
    return row


UX1LX_PRESET = {
    "設定項目": UX1LX_SETTING_LABELS,
    "C5-2": ux1lx_preset_row(15, 15, "HRes2", 40, 70),
    "L11-3": ux1lx_preset_row(4, 5, "HGen", 40, 60),
    "HL18-4": ux1lx_preset_row(4, 3, "HRes", 40, 60),
    "EC9-3": ux1lx_preset_row(
        2.5, 5, "HRes", 40, 60, display_angle="Off"
    ),
    "MC10-3": ux1lx_preset_row(4, 5, "HGen", 40, 60),
    "WL13-3": ux1lx_preset_row(4, 5, "HGen", 35, 60),
    "X20L": ux1lx_preset_row(4, 5, "HRes2", 40, 60),
}


# =========================================================
# HS2：HS2本体条件.xlsx
# =========================================================
HS2_SETTING_LABELS = SETTING_LABELS + ["2D表示角"]


def hs2_preset_row(depth, focus, thi, frequency, bg, dr, display_angle=""):
    row = preset_row(
        "全てセンター", depth, focus, thi, frequency, bg, dr
    )
    row[9] = 50  # コントラスト：全プローブ共通
    row[10] = 1  # グレイマップ：全プローブ共通
    row[20] = 10  # 時間平均
    row.append(display_angle)
    return row


HS2_PRESET = {
    "設定項目": HS2_SETTING_LABELS,
    "C5-2": hs2_preset_row(15, 15, "On", "HRes", 40, 70),
    "L18-4": hs2_preset_row(4, 3, "On", "HRes2", 40, 60),
    "L14-4": hs2_preset_row(4, 3, "On", "HRes", 30, 60),
    "L11-3": hs2_preset_row(4, 5, "On", "HGen", 40, 60),
    "HL18-4": hs2_preset_row(4, 3, "Off", "HRes2", 40, 60),
    "EC9-3": hs2_preset_row(
        2, 5, "On", "HRes", 40, 60, display_angle="Off"
    ),
    "MC10-3": hs2_preset_row(4, 5, "On", "HGen", 40, 60),
    "WL13-3": hs2_preset_row(4, 5, "On", "HGen", 35, 60),
}


# =========================================================
# MX1
# =========================================================
MX1_PRESET = {
    "設定項目": SETTING_LABELS,
    "C5-2": preset_row("全てセンター", 15, 15, "On", "HGen", 40, 70),
    "L18-4": preset_row("全てセンター", 4, 3, "On", "HRes", 40, 60),
    "L14-4": preset_row("全てセンター", 4, 3, "On", "HRes", 30, 60),
    "L11-3": preset_row("全てセンター", 4, 5, "On", "HGen", 40, 60),
    "HL18-4": preset_row("全てセンター", 4, 3, "Off", "HRes", 40, 60),
    "MC10-3": preset_row("全てセンター", 4, 5, "On", "HGen", 40, 60),
}


# =========================================================
# GX30/GX20：GX30本体条件設定.xlsx
# =========================================================
GX30_SETTING_LABELS = SETTING_LABELS + ["2D表示角"]


def gx30_preset_row(depth, focus, dr, display_angle=""):
    row = preset_row(
        "全てセンター", depth, focus, "On", "HRes", 40, dr
    )
    row[9] = 50  # コントラスト：全プローブ共通
    row[10] = 1  # グレイマップ：全プローブ共通
    row[20] = 10  # 時間平均
    row.append(display_angle)
    return row


GX30_PRESET = {
    "設定項目": GX30_SETTING_LABELS,
    "C5-2": gx30_preset_row(15, 15, 70),
    "L18-4": gx30_preset_row(4, 3, 60),
    "EC9-3": gx30_preset_row(3, 5, 60, display_angle="Off"),
}

DEVICE_PRESETS = {
    "UX1/LX": UX1LX_PRESET,
    "HS2": HS2_PRESET,
    "MX1": MX1_PRESET,
    "GX30/GX20": GX30_PRESET,
}


# =========================================================
# リニアプローブのROI
# =========================================================
HS2_ROI_PARAMS = {
    "L11-3": {"roi_x": 176, "roi_y": 74, "roi_width": 498, "roi_height": 60},
    "HL18-4": {"roi_x": 262, "roi_y": 74, "roi_width": 327, "roi_height": 51},
    "WL13-3": {"roi_x": 95, "roi_y": 74, "roi_width": 660, "roi_height": 87},
    "L14-4": {"roi_x": 176, "roi_y": 74, "roi_width": 498, "roi_height": 60},
    "L18-4": {"roi_x": 176, "roi_y": 74, "roi_width": 498, "roi_height": 45},
}

LUXO_ROI_PARAMS = {
    "X20L": {"roi_x": 275, "roi_y": 101, "roi_width": 778, "roi_height": 125},
    "L11-3": {"roi_x": 275, "roi_y": 101, "roi_width": 778, "roi_height": 90},
    "HL18-4": {"roi_x": 407, "roi_y": 101, "roi_width": 514, "roi_height": 83},
    "WL13-3": {"roi_x": 148, "roi_y": 101, "roi_width": 1034, "roi_height": 110},
}


def get_linear_roi_defaults(device_name, probename, img_w, img_h):
    if device_name in ("HS2", "MX1", "GX30/GX20"):
        if probename in HS2_ROI_PARAMS:
            return HS2_ROI_PARAMS[probename].copy()

    if device_name == "UX1/LX" and probename in LUXO_ROI_PARAMS:
        return LUXO_ROI_PARAMS[probename].copy()

    return {
        "roi_x": img_w // 8,
        "roi_y": img_h // 8,
        "roi_width": img_w * 6 // 8,
        "roi_height": img_h // 6,
    }


# =========================================================
# コンベックスプローブの扇形パラメータ
# =========================================================
UX1LX_FAN_PARAMS = {
    "C5-2": (663, -106, 255, 315, -35.0, 35.0),
    "MC10-3": (664, -45, 257, 370, -54.4, 54.4),
    "EC9-3": (425, 73, 53, 85, -88.0, 88.0),
}

MXHS_FAN_PARAMS = {
    "C5-2": (424, -60, 166, 215, -35.0, 35.0),
    "MC10-3": (425, -20, 165, 235, -54.5, 54.5),
    "EC9-3": (425, 71, 54, 85, -86.5, 86.5),
}

DEFAULT_FAN_PARAMS = {
    "C5-2": (663, -106, 255, 320, -35.0, 35.0),
    "MC10-3": (663, -106, 255, 320, -35.0, 35.0),
    "EC9-3": (425, 73, 53, 85, -88.0, 88.0),
}

# EC9-3 の装置別パラメータ
# 出典：EC9-3パラメタ設定_1027.xlsx（シート: HS、B列の装置名に対応）
# 順序：
# (cx, cy, r_inner, r_outer,
#  angle_start, angle_end, angle_bin_size, smooth_window)
EC93_PARAMS = {
    "HS2": (425, 67, 190, 270, -86.5, 86.5, 1, 2),
    "GX30/GX20": (424.5, 71, 144, 265, -85, 85, 1, 2),
    "UX1/LX": (664, 96, 223, 500, -83, 83, 1, 2),
}


def get_fan_params(device_name, probename):
    """扇形パラメータを8要素で返す。"""
    if probename == "EC9-3" and device_name in EC93_PARAMS:
        return EC93_PARAMS[device_name]

    if device_name == "UX1/LX" and probename in UX1LX_FAN_PARAMS:
        base = UX1LX_FAN_PARAMS[probename]
    elif (
        device_name in ("HS2", "MX1", "GX30/GX20")
        and probename in MXHS_FAN_PARAMS
    ):
        base = MXHS_FAN_PARAMS[probename]
    else:
        base = DEFAULT_FAN_PARAMS.get(
            probename,
            (663, -106, 255, 320, -35.0, 35.0),
        )

    return (*base, 0.35, 1)


def make_fan_mask(
    img_h,
    img_w,
    cx,
    cy,
    r_inner,
    r_outer,
    angle_start,
    angle_end,
):
    yy, xx = np.ogrid[:img_h, :img_w]
    dx = xx - cx
    dy = yy - cy

    dist = np.sqrt(dx**2 + dy**2)
    angle_map = np.degrees(np.arctan2(dx, dy))

    return (
        (dist >= r_inner)
        & (dist <= r_outer)
        & (angle_map >= angle_start)
        & (angle_map <= angle_end)
    )


def angle_to_ch(angle_deg, angle_start, angle_end, ch_total, reverse=False):
    """角度をch番号に変換する。reverse=Trueなら右端がch1。"""
    ratio = (angle_deg - angle_start) / (angle_end - angle_start)
    ch_idx = 1 + ratio * (ch_total - 1)
    ch = int(np.round(np.clip(ch_idx, 1, ch_total)))

    # 右端がch1の場合は番号を反転する
    if reverse:
        ch = ch_total + 1 - ch

    return ch


def build_polar_image(
    gray_img,
    cx,
    cy,
    r_inner,
    r_outer,
    angle_start,
    angle_end,
    out_w=512,
    out_h=128,
):
    angles = np.linspace(angle_start, angle_end, out_w)
    radii = np.linspace(r_inner, r_outer, out_h)

    ang_rad = np.radians(angles)
    r_grid = radii[:, np.newaxis]

    img_x = cx + r_grid * np.sin(ang_rad)
    img_y = cy + r_grid * np.cos(ang_rad)

    polar = cv2.remap(
        gray_img.astype(np.float32),
        img_x.astype(np.float32),
        img_y.astype(np.float32),
        interpolation=cv2.INTER_LINEAR,
        borderMode=cv2.BORDER_CONSTANT,
        borderValue=0,
    )
    return polar.astype(np.uint8)


def draw_fan_overlay(
    gray_img,
    fan_mask,
    cx,
    cy,
    r_inner,
    r_outer,
    angle_start,
    angle_end,
):
    vis_img = cv2.cvtColor(gray_img, cv2.COLOR_GRAY2BGR)
    draw_center = (int(round(cx)), int(round(cy)))

    cv_angle_start = int(round(90 - angle_end))
    cv_angle_end = int(round(90 - angle_start))

    cv2.ellipse(
        vis_img,
        draw_center,
        (int(r_outer), int(r_outer)),
        0,
        cv_angle_start,
        cv_angle_end,
        (0, 220, 0),
        2,
    )

    if r_inner > 0:
        cv2.ellipse(
            vis_img,
            draw_center,
            (int(r_inner), int(r_inner)),
            0,
            cv_angle_start,
            cv_angle_end,
            (0, 220, 0),
            2,
        )

    for ang_deg in (angle_start, angle_end):
        ang_rad = np.radians(ang_deg)

        p_inner = (
            int(round(cx + r_inner * np.sin(ang_rad))),
            int(round(cy + r_inner * np.cos(ang_rad))),
        )
        p_outer = (
            int(round(cx + r_outer * np.sin(ang_rad))),
            int(round(cy + r_outer * np.cos(ang_rad))),
        )
        cv2.line(vis_img, p_inner, p_outer, (0, 220, 0), 2)

    cv2.drawMarker(
        vis_img,
        draw_center,
        (0, 0, 255),
        markerType=cv2.MARKER_CROSS,
        markerSize=20,
        thickness=2,
    )

    green_layer = np.zeros_like(vis_img)
    green_layer[fan_mask] = [0, 80, 0]

    return cv2.addWeighted(vis_img, 1.0, green_layer, 0.4, 0)


# =========================================================
# 表示・入力用の共通関数
# =========================================================
def save_analysis_state(idx, data):
    """解析結果をプローブごと(キー末尾に _idx)に保存する。"""
    st.session_state.update(
        {f"{k}_{idx}": v for k, v in data.items()}
    )


def load_uploaded_image(uploaded_file):
    file_bytes = np.frombuffer(uploaded_file.getvalue(), np.uint8)
    raw_img = cv2.imdecode(file_bytes, cv2.IMREAD_COLOR)

    if raw_img is None:
        st.error("画像を読み込めませんでした。PNG、JPG、BMPを確認してください。")
        st.stop()

    gray_img = cv2.cvtColor(raw_img, cv2.COLOR_BGR2GRAY)
    return raw_img, gray_img


def render_preset_table(device_name, probename):
    active_preset = DEVICE_PRESETS.get(device_name, DEFAULT_PRESET)

    if probename not in active_preset:
        active_preset = DEFAULT_PRESET

    preset_df = pd.DataFrame(active_preset)[["設定項目", probename]]

    # 「2D表示角」はHS2・UX1/LX・GX30/GX20のEC9-3のみ表示する。
    if (
        device_name in ("HS2", "UX1/LX", "GX30/GX20")
        and probename != "EC9-3"
    ):
        preset_df = preset_df[
            preset_df["設定項目"] != "2D表示角"
        ]

    table_html = preset_df.to_html(index=False)

    styled_html = f"""
<style>
.preset-table {{
    border-collapse: collapse;
    width: 50%;
    font-size: 20px;
}}
.preset-table th {{
    background-color: #2c3e50;
    color: white;
    padding: 8px 12px;
    border: 2px solid #555;
    text-align: center;
    font-weight: bold;
}}
.preset-table td {{
    padding: 6px 12px;
    border: 2px solid #555;
    text-align: center;
    color: #000;
    font-weight: bold;
}}
.preset-table tr:nth-child(even) {{
    background-color: #f2f2f2;
}}
</style>
{table_html.replace(
    '<table border="1" class="dataframe">',
    '<table class="preset-table">'
)}
"""
    st.markdown(styled_html, unsafe_allow_html=True)


def render_photo_input(title, key_prefix, caption):
    st.header(title)

    method = st.radio(
        "撮影方法",
        options=[
            "選択してください",
            "カメラで撮影する",
            "画像をアップロードする",
        ],
        horizontal=True,
        key=f"{key_prefix}_method",
        label_visibility="collapsed",
    )

    img = None

    if method == "カメラで撮影する":
        cam = st.camera_input(
            f"{caption}を撮影してください",
            key=f"{key_prefix}_camera",
        )
        if cam is not None:
            img = Image.open(cam).convert("RGB")

    elif method == "画像をアップロードする":
        up = st.file_uploader(
            f"{caption}をアップロードしてください (PNG, JPG, BMP)",
            type=["png", "jpg", "jpeg", "bmp"],
            key=f"{key_prefix}_upload",
        )
        if up is not None:
            img = Image.open(up).convert("RGB")

    if img is not None:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.image(img, caption=caption, use_container_width=True)


def render_additional_records(
    key_prefix, inspection_passed, idx, sw_version
):
    """S/N・外観・自由記述欄・プローブ自己診断結果・総合判定・
    PDF保存ボタンを表示する。

    key_prefix: ウィジェットkeyの接頭辞(プローブごとに別)
    inspection_passed: プローブ点検結果が合格ならTrue
    idx: プローブの番号(0始まり)
    sw_version: 共通で選択されたSW Ver
    """
    st.markdown("---")
    render_photo_input(
        "●プローブのS/N",
        f"{key_prefix}_sn",
        "プローブS/N(ラベル)画像",
    )

    st.markdown("---")
    st.header("●プローブ外観")

    for i in range(1, 4):
        st.markdown(f"**📷 外観画像 {i}枚目**")
        render_photo_input(
            "",
            f"{key_prefix}_appearance_{i}",
            f"プローブ外観画像({i}枚目)",
        )

    st.markdown("---")
    now = datetime.datetime.now().replace(microsecond=0)
    st.header(f"●実行日時:{now}")

    st.header("●自由記述欄")
    st.text_area("コメントを入力してください", key=f"{key_prefix}_comment")

    # ---- プローブ自己診断機能(CSV):自由記述欄の下に表示 ----
    # SW Verが「V1.03以降」「V1.10以降」の場合のみ表示する
    if sw_version in DIAGNOSTIC_SW_VERSIONS:
        st.markdown("---")
        has_csv, diagnosis_code = render_diagnostic_csv_section(idx)

        # CSVがアップロードされている場合のみ総合判定を表示する
        # (CSVが削除された場合は非表示になる)
        if has_csv:
            st.markdown("---")
            render_overall_judgement(inspection_passed, diagnosis_code)

    st.markdown("---")
    render_print_button(key_prefix)


# =========================================================
# PDF出力(印刷)用の共通設定
# =========================================================
PRINT_CSS = """
<style>
/* 一括PDF出力時:全タブの内容を表示する */
body.print-all-tabs [role="tabpanel"],
body.print-all-tabs [role="tabpanel"][hidden],
body.print-all-tabs [data-baseweb="tab-panel"] {
    display: block !important;
    visibility: visible !important;
    height: auto !important;
    overflow: visible !important;
}

@media print {
    [data-testid="stSidebar"] { display: none !important; }
    [data-testid="stToolbar"] { display: none !important; }
    [data-testid="stHeader"] { display: none !important; }
    [data-testid="stDecoration"] { display: none !important; }
    iframe { display: none !important; }

    /* ボタン類は印刷しない */
    [data-testid="stButton"],
    [data-testid="stDownloadButton"] { display: none !important; }

    /* 一括PDF出力時はタブの見出しを印刷しない */
    body.print-all-tabs [data-baseweb="tab-list"],
    body.print-all-tabs [data-baseweb="tab-border"],
    body.print-all-tabs [data-baseweb="tab-highlight"] {
        display: none !important;
    }

    /* グラフや画像がページの途中で切れないようにする */
    .stPlotlyChart,
    [data-testid="stImage"] { break-inside: avoid; }
}
</style>
"""

# 全タブをまとめて印刷するJavaScript
# (__NONCE__ は毎回異なる値に置き換え、ボタンを何度押しても動作させる)
PRINT_ALL_TABS_JS = """
<script>
// nonce: __NONCE__
(async function () {
    const win = window.parent;
    const doc = win.document;
    const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

    const tabs = Array.from(doc.querySelectorAll('button[role="tab"]'));
    const panels = Array.from(doc.querySelectorAll('[role="tabpanel"]'));

    if (tabs.length === 0) {
        win.print();
        return;
    }

    const originalIdx = Math.max(
        0,
        tabs.findIndex((t) => t.getAttribute("aria-selected") === "true")
    );

    // 1. 各タブを順に表示して、グラフを正しい幅で描画させる
    for (const tab of tabs) {
        tab.click();
        await sleep(600);
    }

    // 2. 閉じているエキスパンダー(感度低下chの詳細など)を開く
    const closedDetails = Array.from(doc.querySelectorAll("details:not([open])"));
    closedDetails.forEach((d) => { d.open = true; });

    // 3. 全タブを表示し、2本目以降は改ページする
    doc.body.classList.add("print-all-tabs");
    panels.forEach((p, i) => { if (i > 0) p.style.breakBefore = "page"; });
    win.dispatchEvent(new Event("resize"));
    await sleep(1000);

    // 5. 印刷後に元の表示へ戻す
    const cleanup = () => {
        doc.body.classList.remove("print-all-tabs");
        panels.forEach((p) => { p.style.breakBefore = ""; });
        closedDetails.forEach((d) => { d.open = false; });
        win.dispatchEvent(new Event("resize"));
        if (tabs[originalIdx]) tabs[originalIdx].click();
    };
    win.addEventListener("afterprint", cleanup, { once: true });

    // 4. 印刷ダイアログを開く
    win.print();
})();
</script>
"""


def inject_print_css():
    """印刷用CSSをページ全体に適用する。"""
    st.markdown(PRINT_CSS, unsafe_allow_html=True)


def render_sidebar_pdf_all_button(probe_count):
    """サイドバーに、全タブのレポートをまとめてPDF出力するボタンを表示する。"""
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📄 レポート一括PDF出力")

    if probe_count < 2:
        st.sidebar.caption(
            "プローブが2本以上の場合に使用できます。"
            "1本の場合は、画面下部の「この画面をPDFとして保存する」を"
            "使用してください。"
        )
        return

    st.sidebar.info(
        f"💡 全{probe_count}本のプローブのレポートを、"
        "1つのPDFにまとめて出力します。"
        "印刷ダイアログが開いたら、送信先を「PDFに保存」に変更してください。"
    )

    if st.sidebar.button(
        "🖨️ 全プローブのレポートをまとめてPDF出力",
        use_container_width=True,
        type="primary",
        key="sidebar_print_all_tabs",
    ):
        nonce = str(datetime.datetime.now().timestamp())
        with st.sidebar:
            components.html(
                PRINT_ALL_TABS_JS.replace("__NONCE__", nonce),
                height=0,
            )


def render_print_button(key_prefix):
    st.subheader("📄 分析レポートのダウンロード")
    st.info(
        "💡 ボタンを押すとブラウザの印刷ダイアログが開きます。"
        "送信先を「PDFに保存」に変更してください。"
        "(複数本の場合、このボタンでは表示中のタブの内容のみ印刷されます。"
        "全プローブをまとめて出力する場合は、サイドバーの"
        "「全プローブのレポートをまとめてPDF出力」を使用してください)"
    )

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:
        if st.button(
            "🖨️ この画面をPDFとして保存する",
            use_container_width=True,
            type="primary",
            key=f"{key_prefix}_print",
        ):
            nonce = str(datetime.datetime.now().timestamp())
            components.html(
                f"<script>// {nonce}\nwindow.parent.print();</script>",
                height=0,
            )


def download_dataframe(label, df, file_name, key):
    csv_data = df.to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        label=label,
        data=csv_data,
        file_name=file_name,
        mime="text/csv",
        use_container_width=True,
        key=key,
    )


def render_ch_direction_selector(key):
    """ch番号の向き設定を表示し、右端がch1ならTrueを返す。

    感度低下chがある場合に、コンベックス・リニアの両方で使用する。
    """
    st.markdown("---")
    st.subheader("🔄 ch番号の向き設定")

    ch_direction = st.radio(
        "ch1はどちら側ですか？",
        options=["左端がch1", "右端がch1"],
        index=0,
        horizontal=True,
        key=key,
    )
    return ch_direction == "右端がch1"


# =========================================================
# コンベックスプローブの解析
# =========================================================
def render_convex_analysis(
    uploaded_file, device_name, probename, idx, sw_version
):
    if uploaded_file is None:
        st.info(
            f"👈 サイドバーからプローブ{idx + 1}の"
            "超音波画像をアップロードしてください。"
        )
        return

    raw_img, gray_img = load_uploaded_image(uploaded_file)
    img_h, img_w = gray_img.shape

    if probename == "C5-2":
        chnumber = 160
    elif probename == "EC9-3":
        chnumber = 192
    else:
        chnumber = 128

    (
        cx,
        cy,
        r_inner,
        r_outer,
        angle_start,
        angle_end,
        angle_bin_size,
        smooth_window,
    ) = get_fan_params(device_name, probename)

    fan_mask = make_fan_mask(
        img_h,
        img_w,
        cx,
        cy,
        r_inner,
        r_outer,
        angle_start,
        angle_end,
    )

    if not np.any(fan_mask):
        st.error("扇形マスク内に有効なピクセルがありません。")
        return

    vis_img_overlay = draw_fan_overlay(
        gray_img,
        fan_mask,
        cx,
        cy,
        r_inner,
        r_outer,
        angle_start,
        angle_end,
    )

    polar_w = 512
    polar_h = 128

    polar_img = build_polar_image(
        gray_img,
        cx,
        cy,
        r_inner,
        r_outer,
        angle_start,
        angle_end,
        polar_w,
        polar_h,
    )

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("🖼️ 元画像（扇形スキャンエリア表示）")
        st.image(
            cv2.cvtColor(vis_img_overlay, cv2.COLOR_BGR2RGB),
            use_container_width=True,
        )

    with col2:
        st.subheader("🖼️ 極座標変換画像（扇形 → 矩形）")
        st.image(polar_img, use_container_width=True)

    valid_mask_polar = polar_img > 0

    col_sum = np.sum(
        polar_img.astype(np.float32) * valid_mask_polar,
        axis=0,
    )
    col_count = np.sum(valid_mask_polar, axis=0).astype(np.float32)
    col_count[col_count == 0] = 1.0

    mean_profile_polar = col_sum / col_count

    if smooth_window > 1:
        mean_profile_polar = (
            pd.Series(mean_profile_polar)
            .rolling(
                window=int(smooth_window),
                center=True,
                min_periods=1,
            )
            .mean()
            .to_numpy()
        )

    angle_axis = np.linspace(angle_start, angle_end, polar_w)
    angle_binned = (
        np.round(angle_axis / angle_bin_size) * angle_bin_size
    )

    df_polar = pd.DataFrame(
        {
            "角度ビン(°)": np.round(angle_binned, 2),
            "平均輝度": mean_profile_polar,
        }
    )

    df_angle_stats = (
        df_polar.groupby("角度ビン(°)", as_index=False)["平均輝度"]
        .mean()
        .round(2)
    )

    df_angle_stats_full = (
        df_polar.groupby("角度ビン(°)")["平均輝度"]
        .agg(
            ピクセル数="count",
            平均輝度="mean",
            最大輝度="max",
            最小輝度="min",
            標準偏差="std",
        )
        .round(2)
        .reset_index()
    )

    st.markdown("---")
    st.subheader("📈 輝度値のプロファイル")

    fig_angle = px.line(
        df_angle_stats,
        x="角度ビン(°)",
        y="平均輝度",
        labels={
            "角度ビン(°)": "角度 (°)",
            "平均輝度": "平均輝度値 (0-255)",
        },
        title="角度ごとの平均輝度プロファイル（極座標変換）",
        markers=True,
    )

    fig_angle.update_traces(
        line_color="#00A8FF",
        marker=dict(size=4),
    )
    fig_angle.update_layout(
        height=280,
        yaxis_range=[0, 255],
        margin=dict(l=20, r=20, t=40, b=20),
        yaxis=dict(tickmode="linear", tick0=0, dtick=5),
        xaxis=dict(ticksuffix="°"),
    )
    st.plotly_chart(
        fig_angle,
        use_container_width=True,
        key=f"convex_angle_chart_{idx}",
    )

    download_dataframe(
        "📥 輝度プロファイルデータ (CSV) をダウンロード",
        df_angle_stats_full,
        f"angle_profile_{uploaded_file.name}.csv",
        f"convex_angle_csv_{idx}",
    )

    brightness_vals = df_angle_stats["平均輝度"].to_numpy()

    mean_profile2 = float(np.mean(brightness_vals))
    min_val2 = int(np.min(brightness_vals))

    normal_mask = np.abs(brightness_vals - mean_profile2) < 20
    new_mean = (
        float(np.mean(brightness_vals[normal_mask]))
        if normal_mask.any()
        else mean_profile2
    )

    mean_profile3 = new_mean - brightness_vals
    meanmin_val3 = new_mean - min_val2
    diff_values = -mean_profile3

    fig_diff = px.line(
        x=df_angle_stats["角度ビン(°)"],
        y=diff_values,
        labels={
            "x": "角度 (°)",
            "y": "輝度差分 (0-255)",
        },
        title="ROI領域における平均輝度値との差分",
    )
    fig_diff.add_hline(
        y=-20,
        line_dash="dash",
        line_color="red",
        line_width=2,
        annotation_text="閾値: -20（感度低下）",
    )
    fig_diff.update_traces(line_color="#00A8FF")
    fig_diff.update_layout(
        height=280,
        margin=dict(l=20, r=20, t=30, b=20),
        yaxis=dict(tickmode="linear", tick0=0, dtick=5),
    )
    st.plotly_chart(
        fig_diff,
        use_container_width=True,
        key=f"convex_diff_chart_{idx}",
    )

    profile_df2 = pd.DataFrame(
        {
            "角度(°)": df_angle_stats["角度ビン(°)"].to_numpy(),
            "Y座標_輝度値": diff_values,
        }
    )
    download_dataframe(
        "📥 平均輝度差分プロファイルデータ (CSV) をダウンロード",
        profile_df2,
        f"diff_profile_{uploaded_file.name}.csv",
        f"convex_diff_csv_{idx}",
    )

    bad_bin_mask = mean_profile3 > 20
    bad_angles = df_angle_stats["角度ビン(°)"].to_numpy()[bad_bin_mask]

    max_drop = (
        float(np.max(mean_profile3))
        if len(mean_profile3) > 0
        else 0.0
    )

    # 感度低下がある場合のみ、ch番号の向き設定を表示
    if max_drop > 20:
        ch_reverse = render_ch_direction_selector(
            f"ch_direction_radio_convex_{idx}"
        )
    else:
        ch_reverse = False

    bad_channels = sorted(
        {
            angle_to_ch(
                angle,
                angle_start,
                angle_end,
                chnumber,
                reverse=ch_reverse,
            )
            for angle in bad_angles
        }
    )

    st.markdown("---")
    st.header("📊 プローブ点検結果")

    # プローブ点検結果の合否(総合判定で使用)
    inspection_passed = not (max_drop > 20)

    if not inspection_passed:
        st.error("プローブ点検：感度低下chあり🚨")

        with st.expander("感度低下している詳細chを確認する"):
            if bad_channels:
                for ch in bad_channels:
                    st.write(f"{ch}ch")
            else:
                st.write("該当chなし")
    else:
        st.success("プローブ点検：合格🎉")

    if max_drop > 20 and len(bad_angles) > 0:
        img_rgb = cv2.cvtColor(raw_img, cv2.COLOR_BGR2RGB)
        overlay_img = img_rgb.copy()

        drop_per_col = new_mean - mean_profile_polar
        drop_per_col_clipped = np.clip(drop_per_col, 0, None)

        max_possible = max(
            float(np.max(drop_per_col_clipped)),
            1.0,
        )
        heat_intensity = np.clip(
            drop_per_col_clipped / max_possible * 255.0,
            0,
            255,
        ).astype(np.uint8)

        bad_col_mask = drop_per_col > 20

        polar_heat_gray = np.zeros(
            (polar_h, polar_w),
            dtype=np.uint8,
        )
        polar_heat_gray[:, bad_col_mask] = heat_intensity[bad_col_mask]

        valid_col_polar = np.sum(valid_mask_polar, axis=0) > 0
        polar_heat_gray[:, ~valid_col_polar] = 0

        polar_heat_color = cv2.applyColorMap(
            255 - polar_heat_gray,
            cv2.COLORMAP_JET,
        )
        polar_heat_rgb = cv2.cvtColor(
            polar_heat_color,
            cv2.COLOR_BGR2RGB,
        )

        yy, xx = np.ogrid[:img_h, :img_w]

        dx_all = (xx - cx).astype(np.float32)
        dy_all = (yy - cy).astype(np.float32)

        ang_all = np.degrees(np.arctan2(dx_all, dy_all))
        dist_all = np.sqrt(dx_all**2 + dy_all**2)

        col_idx = np.clip(
            np.round(
                (ang_all - angle_start)
                / (angle_end - angle_start)
                * (polar_w - 1)
            ).astype(np.int32),
            0,
            polar_w - 1,
        )

        row_idx = np.clip(
            np.round(
                (dist_all - r_inner)
                / (r_outer - r_inner)
                * (polar_h - 1)
            ).astype(np.int32),
            0,
            polar_h - 1,
        )

        heat_on_orig = np.zeros(
            (img_h, img_w, 3),
            dtype=np.uint8,
        )

        fy, fx = np.where(fan_mask)
        heat_on_orig[fy, fx] = polar_heat_rgb[
            row_idx[fy, fx],
            col_idx[fy, fx],
        ]

        bad_pixel_mask = fan_mask & bad_col_mask[col_idx]
        alpha = 0.7

        overlay_img = np.where(
            bad_pixel_mask[:, :, np.newaxis],
            (
                overlay_img.astype(np.float32) * (1.0 - alpha)
                + heat_on_orig.astype(np.float32) * alpha
            ).astype(np.uint8),
            overlay_img,
        )

        col1, col2 = st.columns(2)

        with col1:
            st.subheader("元画像")
            st.image(img_rgb, use_container_width=True)

        with col2:
            st.subheader("重ね合わせ画像（感度低下ch可視化）")
            st.image(overlay_img, use_container_width=True)

    save_analysis_state(
        idx,
        {
            "analysis_done": True,
            "uploaded_file_name": uploaded_file.name,
            "meanmin_val3": meanmin_val3,
            "mean_profile2": mean_profile2,
            "max_drop": max_drop,
            "min_val2": min_val2,
            "bad_channels": bad_channels,
            "ch_reverse": ch_reverse,
            "chnumber": chnumber,
            "df_angle_stats": df_angle_stats,
            "mean_profile3": mean_profile3,
            "vis_img_overlay": vis_img_overlay,
            "polar_img": polar_img,
            "cx": cx,
            "cy": cy,
            "r_inner": r_inner,
            "r_outer": r_outer,
            "angle_start": angle_start,
            "angle_end": angle_end,
            "angle_bin_size": angle_bin_size,
            "smooth_window": smooth_window,
            "inspection_passed": inspection_passed,
        },
    )

    render_additional_records(
        f"convex_{idx}", inspection_passed, idx, sw_version
    )


# =========================================================
# リニアプローブの解析
# =========================================================
def px_to_ch(px_val, roi_width, chnumber, reverse=False):
    if reverse:
        ch = px_val * (-(chnumber - 1) / roi_width) + chnumber
    else:
        ch = px_val * ((chnumber - 1) / roi_width) + 1
    return round(ch)


def render_linear_analysis(
    uploaded_file, device_name, probename, idx, sw_version
):
    if uploaded_file is None:
        st.info(
            f"👈 サイドバーからプローブ{idx + 1}の"
            "超音波画像をアップロードしてください。"
        )
        return

    if probename in ("X20L", "L18-4", "L11-3", "WL13-3", "L14-4"):
        chnumber = 192
    elif probename in ("HL18-4", "MC10-3"):
        chnumber = 128
    else:
        chnumber = 0.1

    raw_img, gray_img = load_uploaded_image(uploaded_file)
    img_h, img_w = gray_img.shape

    roi_defaults = get_linear_roi_defaults(
        device_name,
        probename,
        img_w,
        img_h,
    )

    roi_x = roi_defaults["roi_x"]
    roi_y = roi_defaults["roi_y"]
    roi_width = roi_defaults["roi_width"]
    roi_height = roi_defaults["roi_height"]

    x_min = max(0, roi_x)
    y_min = max(0, roi_y)
    x_max = min(roi_x + roi_width, img_w)
    y_max = min(roi_y + roi_height, img_h)

    rw = x_max - x_min
    rh = y_max - y_min

    if rw <= 0 or rh <= 0:
        st.error("ROIが画像の範囲外です。")
        return

    roi_crop = gray_img[y_min:y_max, x_min:x_max]

    vis_img = cv2.cvtColor(gray_img, cv2.COLOR_GRAY2BGR)
    cv2.rectangle(
        vis_img,
        (x_min, y_min),
        (x_max, y_max),
        (0, 255, 255),
        3,
    )
    cv2.putText(
        vis_img,
        f"ROI: {rw}x{rh}",
        (x_min, max(y_min - 10, 20)),
        cv2.FONT_HERSHEY_SIMPLEX,
        0.8,
        (0, 255, 255),
        2,
    )

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("🖼️ 元画像")
        st.image(cv2.cvtColor(vis_img, cv2.COLOR_BGR2RGB))

    with col2:
        st.subheader("🖼️ ROI切り出し範囲")
        st.image(roi_crop)

    mean_profile = np.mean(roi_crop, axis=0)
    roix = np.arange(rw)

    # 元コードと同じく左右端の輝度値を補正
    mean_profile[:4] += 5
    mean_profile[max(0, rw - 4):rw] += 5

    max_val2 = int(np.max(mean_profile))
    min_val2 = int(np.min(mean_profile))
    mean_profile2 = float(np.mean(mean_profile))

    st.markdown("---")
    st.subheader("📈 輝度値のプロファイル")

    fig_line = px.line(
        x=roix,
        y=mean_profile,
        labels={
            "x": "X座標 (px)",
            "y": "輝度値 (0-255)",
        },
        title="ROI領域における輝度値",
    )
    fig_line.update_traces(line_color="#00A8FF")
    fig_line.update_layout(
        height=280,
        margin=dict(l=20, r=20, t=30, b=20),
        yaxis=dict(tickmode="linear", tick0=0, dtick=5),
    )
    st.plotly_chart(
        fig_line,
        use_container_width=True,
        key=f"linear_line_chart_{idx}",
    )

    profile_df = pd.DataFrame(
        {
            "X座標_px": roix,
            "Y座標_px": mean_profile,
        }
    )

    download_dataframe(
        "📥 輝度プロファイルデータ (CSV) をダウンロード",
        profile_df,
        f"line_profile_Y{rw}_{uploaded_file.name}.csv",
        f"linear_profile_csv_{idx}",
    )

    threshold2 = 20

    profile_df["差分"] = (
        profile_df["Y座標_px"] - mean_profile2
    ).abs()

    filtered_df = profile_df[profile_df["差分"] < threshold2]

    new_mean = (
        float(filtered_df["Y座標_px"].mean())
        if not filtered_df.empty
        else mean_profile2
    )

    mean_profile3 = new_mean - mean_profile
    meanmin_val3 = new_mean - min_val2

    fig_diff = px.line(
        x=roix,
        y=-mean_profile3,
        labels={
            "x": "X座標 (px)",
            "y": "輝度値 (0-255)",
        },
        title="ROI領域における平均輝度値との差分",
    )
    fig_diff.add_hline(
        y=-20,
        line_dash="dash",
        line_color="red",
        line_width=2,
        annotation_text="閾値: -20（感度低下）",
    )
    fig_diff.update_traces(line_color="#00A8FF")
    fig_diff.update_layout(
        height=280,
        margin=dict(l=20, r=20, t=30, b=20),
        yaxis=dict(tickmode="linear", tick0=0, dtick=5),
    )
    st.plotly_chart(
        fig_diff,
        use_container_width=True,
        key=f"linear_diff_chart_{idx}",
    )

    profile_df2 = pd.DataFrame(
        {
            "X座標_px": roix,
            "Y座標_輝度値": -mean_profile3,
        }
    )

    download_dataframe(
        "📥 平均輝度差分プロファイルデータ (CSV) をダウンロード",
        profile_df2,
        f"diff_profile_{uploaded_file.name}.csv",
        f"linear_diff_csv_{idx}",
    )

    # 感度低下の有無(平均との差分、または-20未満の点があるか)
    has_drop = (
        meanmin_val3 > 20
        or bool((profile_df2["Y座標_輝度値"] < -20).any())
    )

    # 感度低下がある場合のみ、ch番号の向き設定を表示
    if has_drop:
        ch_reverse = render_ch_direction_selector(
            f"ch_direction_radio_linear_{idx}"
        )
    else:
        ch_reverse = False

    bad_profile = profile_df2[
        profile_df2["Y座標_輝度値"] < -20
    ].copy()

    bad_profile["ch番号"] = bad_profile["X座標_px"].apply(
        lambda px_val: px_to_ch(
            px_val,
            rw,
            chnumber,
            reverse=ch_reverse,
        )
    )

    unique_channels = sorted(set(bad_profile["ch番号"].tolist()))
    unique_data2 = [f"{ch}ch" for ch in unique_channels]

    st.markdown("---")
    st.header("📊 プローブ点検結果")

    # プローブ点検結果の合否(総合判定で使用)
    inspection_passed = not has_drop

    if not inspection_passed:
        st.error("プローブ点検：感度低下chあり🚨")

        with st.expander("感度低下している詳細chを確認する"):
            if unique_data2:
                for ch in unique_data2:
                    st.write(ch)
            else:
                st.write("該当chなし")
    else:
        st.success("プローブ点検：合格🎉")

    if has_drop and not bad_profile.empty:
        img_rgb = cv2.cvtColor(raw_img, cv2.COLOR_BGR2RGB)
        overlay_img = img_rgb.copy()

        roi_rgb = img_rgb[y_min:y_max, x_min:x_max].copy()
        gray_roi = cv2.cvtColor(roi_rgb, cv2.COLOR_RGB2GRAY)

        bad_columns = bad_profile["X座標_px"].astype(int).tolist()
        selected_gray = gray_roi[:, bad_columns]

        if new_mean != 0:
            norm_gray = np.clip(
                (new_mean - selected_gray) / new_mean * 255.0,
                0,
                255,
            ).astype(np.uint8)
        else:
            norm_gray = np.zeros_like(
                selected_gray,
                dtype=np.uint8,
            )

        heatmap_bgr = cv2.applyColorMap(
            255 - norm_gray,
            cv2.COLORMAP_JET,
        )
        heatmap_rgb = cv2.cvtColor(
            heatmap_bgr,
            cv2.COLOR_BGR2RGB,
        )

        heatmap_roi_full = np.zeros(
            (rh, rw, 3),
            dtype=np.uint8,
        )

        for hidx, col_px in enumerate(bad_columns):
            if 0 <= col_px < rw:
                heatmap_roi_full[:, col_px, :] = heatmap_rgb[:, hidx, :]

        roi_base = overlay_img[
            y_min:y_max,
            x_min:x_max,
        ].astype(np.float32)

        heat_mask = np.any(
            heatmap_roi_full > 0,
            axis=2,
            keepdims=True,
        )

        blended_roi = np.where(
            heat_mask,
            (
                roi_base * 0.3
                + heatmap_roi_full.astype(np.float32) * 0.7
            ).astype(np.uint8),
            roi_base.astype(np.uint8),
        )

        overlay_img[y_min:y_max, x_min:x_max] = blended_roi

        col1, col2 = st.columns(2)

        with col1:
            st.subheader("元画像")
            st.image(img_rgb, use_container_width=True)

        with col2:
            st.subheader("重ね合わせ画像")
            st.image(overlay_img, use_container_width=True)

    save_analysis_state(
        idx,
        {
            "analysis_done": True,
            "uploaded_file_name": uploaded_file.name,
            "meanmin_val3": meanmin_val3,
            "mean_profile2": mean_profile2,
            "max_val2": max_val2,
            "min_val2": min_val2,
            "unique_data2": unique_data2,
            "ch_reverse": ch_reverse,
            "chnumber": chnumber,
            "roix": roix,
            "mean_profile": mean_profile,
            "mean_profile3": mean_profile3,
            "vis2": vis_img,
            "roi_crop": roi_crop,
            "inspection_passed": inspection_passed,
        },
    )

    render_additional_records(
        f"linear_{idx}", inspection_passed, idx, sw_version
    )


# =========================================================
# ホーム画面
# =========================================================
def render_home():
    st.subheader("●点検するプローブの本数の選択")

    probe_count = st.radio(
        "点検するプローブの本数",
        options=PROBE_COUNT_OPTIONS,
        index=0,
        horizontal=True,
        format_func=lambda n: f"{n}本",
        key="probe_count",
        label_visibility="collapsed",
    )

    st.markdown("---")

    # ---- 超音波診断装置本体・SW Verは全プローブ共通で1回だけ選択 ----
    device_name, sw_version = render_common_selection()

    st.markdown("---")

    if probe_count == 1:
        # 1本の場合はタブを使わず、そのまま表示する
        render_probe_panel(0, probe_count, device_name, sw_version)
    else:
        # 2本・3本の場合は、プローブ名の選択以降をタブで切り替える
        tab_labels = [
            f"🔹 プローブ{i + 1}" for i in range(probe_count)
        ]
        tabs = st.tabs(tab_labels)

        for idx, tab in enumerate(tabs):
            with tab:
                render_probe_panel(idx, probe_count, device_name, sw_version)

    # ---- サイドバー:全プローブのレポートを一括でPDF出力 ----
    render_sidebar_pdf_all_button(probe_count)


def render_probe_panel(idx, probe_count, device_name, sw_version):
    """1本分の「プローブ名の選択」以降の表示をまとめて行う。"""
    st.header(f"🔹 プローブ {idx + 1} / {probe_count}")

    probename, uploaded_file = render_selection(
        idx, device_name, sw_version
    )

    render_probe_inspection(
        idx, probename, device_name, sw_version, uploaded_file
    )


# ---------------------------------------------------------
# 共通選択欄(超音波診断装置本体・SW Ver):全プローブで1回のみ
# ---------------------------------------------------------
def render_common_selection():
    """超音波診断装置本体とSW Verを選択する(全プローブ共通)。

    戻り値: (device_name, sw_version)
    SW Ver選択が必要な装置でSW Ver未選択の場合、sw_versionは""。
    """
    st.subheader("●超音波診断装置本体の選択")

    device_choices = [""] + DEVICE_OPTIONS
    saved_device = st.session_state.get("common_device_name", "")
    device_index = (
        device_choices.index(saved_device)
        if saved_device in device_choices
        else 0
    )

    device_name = st.selectbox(
        label="超音波診断装置本体",
        options=device_choices,
        index=device_index,
        key="common_device_selector",
        label_visibility="collapsed",
    )
    st.session_state["common_device_name"] = device_name

    # ---- SW Verの選択(UX1/LX・MX1のみ) ----
    sw_options = SW_VERSION_MAP.get(device_name)
    sw_version = ""

    if sw_options:
        st.subheader("●SW Verの選択")

        sw_choices = [""] + sw_options
        saved_sw = st.session_state.get("common_sw_version", "")
        sw_index = (
            sw_choices.index(saved_sw)
            if saved_sw in sw_choices
            else 0
        )

        sw_version = st.selectbox(
            label="SW Ver",
            options=sw_choices,
            index=sw_index,
            key=f"common_sw_selector_{device_name.replace('/', '_')}",
            label_visibility="collapsed",
        )

    st.session_state["common_sw_version"] = sw_version

    return device_name, sw_version


# ---------------------------------------------------------
# プローブごとの選択欄(プローブ名のみ)
# ---------------------------------------------------------
def render_selection(idx, device_name, sw_version):
    n = idx + 1

    sw_options = SW_VERSION_MAP.get(device_name)
    sw_ok = (not sw_options) or bool(sw_version)
    device_ready = bool(device_name) and sw_ok

    st.subheader(f"●プローブ{n}:プローブ名の選択")

    if not device_ready:
        # 装置本体(およびSW Ver)が未選択の間はプローブ名を選べない
        st.info(
            "👆 先に「超音波診断装置本体」"
            "(およびSW Ver)を選択してください。"
        )
        st.sidebar.warning(
            f"⚠️ プローブ{n}:超音波診断装置本体(およびSW Ver)と"
            "プローブ名を選択すると画像をアップロードできます。"
        )
        return "", None

    # 選択した装置に対応するプローブのみ表示する
    probe_choices = [""] + [
        p
        for p in PROBE_OPTIONS
        if p and device_name in PROBE_DEVICE_MAP.get(p, [])
    ]

    saved_probe = st.session_state.get(f"probename_{idx}", "")
    probe_index = (
        probe_choices.index(saved_probe)
        if saved_probe in probe_choices
        else 0
    )

    probename = st.selectbox(
        label=f"プローブ{n} 名",
        options=probe_choices,
        index=probe_index,
        # 装置を変更した場合に選択状態をリセットするため、keyに装置名を含める
        key=f"probe_selector_{idx}_{device_name.replace('/', '_')}",
        label_visibility="collapsed",
    )
    st.session_state[f"probename_{idx}"] = probename

    if probename:
        uploaded_file = st.sidebar.file_uploader(
            f"【プローブ{n}】超音波画像をアップロード (PNG, JPG, BMP)",
            type=["png", "jpg", "jpeg", "bmp"],
            key=f"main_uploader_{idx}",
        )
    else:
        uploaded_file = None
        st.sidebar.warning(
            f"⚠️ プローブ{n}:プローブ名を選択すると"
            "画像をアップロードできます。"
        )

    return probename, uploaded_file


# ---------------------------------------------------------
# プローブ点検(設定項目表示・画像解析)
# ---------------------------------------------------------
def render_probe_inspection(
    idx, probename, device_name, sw_version, uploaded_file
):
    if not probename:
        st.markdown(
            """
<div style="
    background-color:#fff3cd;
    border:1px solid #ffc107;
    border-radius:5px;
    padding:16px 20px;
    font-size:18px;
    color:#000;
    font-weight:bold;
">
⚠️ 【空中放射画像取得における注意点】<br><br>
・超音波ゲルが音響レンズに付着していないことを確認してください。<br><br>
・ゲルが固着している場合は拭き取ってください。<br><br>
・プローブホルダーなど安定した場所で画像を取得してください。<br><br>
・画像取得時はレンズ表面に触れないでください。<br><br>
・プローブカバーや穿刺ブラケットを外してください。<br><br>
・スキャンエリア内に色付けやテキストを入れないでください。
</div>
""",
            unsafe_allow_html=True,
        )
        return

    st.subheader("●超音波診断装置本体の設定項目")
    if sw_version:
        st.caption(f"装置: {device_name}　/　SW Ver: {sw_version}")
    render_preset_table(device_name, probename)

    st.markdown("---")

    if probename in CONVEX_PROBES:
        render_convex_analysis(
            uploaded_file, device_name, probename, idx, sw_version
        )
    else:
        render_linear_analysis(
            uploaded_file, device_name, probename, idx, sw_version
        )


# ---------------------------------------------------------
# プローブ自己診断機能（CSV）
# ---------------------------------------------------------
def judge_self_diagnosis(df_raw):
    """CSVのB11セルからプローブ自己診断結果コードを取得する。

    B11 = 0 → 0(合格)
    B11 = 1 → 1(不合格)
    B11 = 2 → 2(プローブ接触不良)
    それ以外・取得不可 → None(判定不能)

    header=None で読み込んでいるため、B11は iloc[10, 1]。
    """
    if df_raw.shape[0] <= 10 or df_raw.shape[1] <= 1:
        return None

    value = pd.to_numeric(df_raw.iloc[10, 1], errors="coerce")

    if pd.isna(value):
        return None

    if value in (0, 1, 2):
        return int(value)

    return None


def show_self_diagnosis_result(code):
    """プローブ自己診断結果コードに応じたメッセージを表示する。"""
    if code == 0:
        st.success("プローブ自己診断機能の結果：合格🎉")
    elif code == 1:
        st.error("プローブ自己診断機能の結果：不合格🚨")
    elif code == 2:
        st.warning("プローブ自己診断機能の結果：プローブ接触不良⚠️")
    else:
        st.info(
            "B11セルが0・1・2のいずれでもないため、"
            "プローブ自己診断結果を判定できません。"
        )


def render_diagnostic_graph(df_raw, idx):
    """CSVのB列(Channel)・F列(Abs. Sensitivity)をグラフ表示する。"""
    if df_raw.shape[0] <= 14 or df_raw.shape[1] <= 5:
        st.info("グラフ表示には15行・F列以上のCSVが必要です。")
        return

    df_plot = df_raw.iloc[14:, [1, 5]].copy()
    df_plot.columns = ["Channel", "Abs. Sensitivity [dB]"]

    for column in df_plot.columns:
        df_plot[column] = pd.to_numeric(df_plot[column], errors="coerce")

    df_plot = df_plot.dropna()

    if df_plot.empty:
        st.info("B列・F列の15行目以降に数値データがありません。")
        return

    fig_csv = go.Figure()
    fig_csv.add_trace(
        go.Scatter(
            x=df_plot["Channel"],
            y=df_plot["Abs. Sensitivity [dB]"],
            mode="lines+markers",
            line=dict(color="#00A8FF", width=2),
            marker=dict(size=5),
            name="Abs. Sensitivity [dB]",
        )
    )
    fig_csv.update_layout(
        height=450,
        title="Absolute Sensitivity",
        xaxis_title="Channel",
        yaxis_title="Absolute Sensitivity [dB]",
        margin=dict(l=20, r=20, t=50, b=20),
        xaxis=dict(
            tickmode="linear",
            dtick=max(1, len(df_plot) // 20),
        ),
        hovermode="x unified",
    )
    st.plotly_chart(
        fig_csv,
        use_container_width=True,
        key=f"csv_chart_{idx}",
    )


def render_diagnostic_csv_section(idx):
    """プローブ自己診断機能のCSV(1ファイルのみ)のアップロード・
    結果表示・グラフ表示。

    アップロード欄の状態に合わせて結果を表示・非表示にする。
    (CSVが削除された場合は、結果・グラフも表示しない)

    戻り値: (has_csv, diagnosis_code)
      has_csv: CSVがアップロードされていればTrue
      diagnosis_code: 0:合格 / 1:不合格 / 2:接触不良 / None:判定不能
    """
    st.header("📁 プローブ自己診断機能の結果表示")
    st.markdown("### 📤 csvファイルのアップロード")

    uploaded_csv = st.file_uploader(
        "CSVファイルを1つ選択してください",
        type=["csv"],
        accept_multiple_files=False,  # 1ファイルのみ
        key=f"diagnostic_csv_uploader_{idx}",
    )

    # ファイルが無い(削除された)場合は、結果・グラフを表示しない
    if uploaded_csv is None:
        st.info("まだファイルがアップロードされていません。")
        return False, None

    st.success(f"✅ 「{uploaded_csv.name}」を受け付けました。")

    st.markdown("---")
    st.markdown("### 📊 プローブ自己診断結果・グラフ")

    try:
        # skip_blank_lines=False：空行を飛ばさず、
        # Excel上の行番号(B11など)とずれないようにする
        df_raw = pd.read_csv(
            io.BytesIO(uploaded_csv.getvalue()),
            header=None,
            encoding="utf-8",
            on_bad_lines="skip",
            skip_blank_lines=False,
        )
    except Exception as exc:
        st.warning(f"CSVの処理中にエラーが発生しました: {exc}")
        return True, None

    # ---- プローブ自己診断結果の判定(B11セル) ----
    diagnosis_code = judge_self_diagnosis(df_raw)
    show_self_diagnosis_result(diagnosis_code)

    # ---- グラフ表示 ----
    try:
        render_diagnostic_graph(df_raw, idx)
    except Exception as exc:
        st.warning(f"グラフの表示中にエラーが発生しました: {exc}")

    return True, diagnosis_code


# ---------------------------------------------------------
# 総合判定
# ---------------------------------------------------------
def render_overall_judgement(inspection_passed, diagnosis_code):
    """プローブ点検結果とプローブ自己診断機能の結果から
    総合判定を表示する。

    - どちらか一方でも不合格 → 不合格
    - 両方とも合格 → 合格
    - 上記以外(接触不良、判定不能) → 判定できません
    """
    st.header("🏁 総合判定")

    if (not inspection_passed) or diagnosis_code == 1:
        st.error("総合判定結果：不合格🚨")
    elif diagnosis_code == 0:
        st.success("総合判定結果：合格🎉")
    elif diagnosis_code == 2:
        st.warning(
            "総合判定結果：判定できません"
            "(プローブ接触不良のため、"
            "プローブ自己診断機能の結果が確定していません)"
        )
    else:
        st.info(
            "総合判定結果：判定できません"
            "(プローブ自己診断結果を判定できません)"
        )


# =========================================================
# アプリについて
# =========================================================
def render_about_page():
    st.subheader("ℹ️ アプリについて")
    st.markdown("---")

    st.markdown(
        """
<div style="
    background-color:#eaf4fb;
    border:1px solid #3498db;
    border-radius:8px;
    padding:20px 24px;
    font-size:16px;
    color:#000;
">
<h3>🔊 超音波プローブ点検アプリ</h3>

<p>
空中放射画像を解析し、チャンネル単位での
感度低下箇所を判定するツールです。
</p>

<hr>

<h4>📌 主な機能</h4>
<ul>
<li>点検するプローブの本数(1〜3本)の選択
(2本以上の場合はタブで切り替え)</li>
<li>超音波診断装置本体・SW Verの共通選択
(複数本でも1回の選択で全プローブに適用)</li>
<li>選択した装置に対応するプローブ名の選択</li>
<li>装置の設定項目の確認</li>
<li>空中放射画像のアップロードとROI解析</li>
<li>輝度プロファイル・差分グラフ表示</li>
<li>感度低下チャンネルの判定</li>
<li>感度低下chがある場合のch番号の向き設定
(リニア・コンベックス共通)</li>
<li>ヒートマップによる可視化</li>
<li>S/Nラベル・プローブ外観画像の記録</li>
<li>プローブ自己診断結果CSV(1ファイル)の判定
(合格/不合格/プローブ接触不良)とグラフ表示
(UX1/LX V1.03以降、MX1 V1.10以降)</li>
<li>プローブ点検結果とプローブ自己診断結果による総合判定
(UX1/LX V1.03以降、MX1 V1.10以降)</li>
<li>ブラウザの印刷機能によるPDF保存</li>
<li>複数本の場合、サイドバーから全プローブのレポートを
1つのPDFにまとめて出力</li>
</ul>

<hr>

<h4>📋 対応プローブ</h4>
<p>
X20L / L18-4 / L11-3 / HL18-4 / WL13-3 /
L14-4 / MC10-3 / EC9-3 / C5-2
</p>

<h4>🖥️ 対応超音波診断装置</h4>
<p>HS2 / MX1 / GX30/GX20 / UX1/LX</p>

<hr>

<p>バージョン: <strong>1.03</strong></p>
</div>
""",
        unsafe_allow_html=True,
    )

    st.markdown("---")
    st.subheader("📖 アプリの使用手順")

    st.markdown(
        """
0. 最初に点検するプローブの本数(1〜3本)を選択します。
   2本・3本の場合は「プローブ1」「プローブ2」…のタブが表示されるので、
   タブを切り替えてプローブごとに操作します。
1. 指定された装置設定でプローブの空中放射画像を取得します。
2. 「ホーム / プローブ選択」ページで、
   超音波診断装置本体を選択します(全プローブ共通で1回のみ)。
   UX1/LX・MX1の場合は、SW Verも選択します。
3. 各プローブのタブで、プローブ名を選択します。
   選択した装置に対応するプローブのみ表示されます。
4. サイドバーから取得した画像をアップロードします。
   (複数本の場合は「【プローブ1】」「【プローブ2】」…と区別されています)
5. 輝度プロファイルと点検結果を確認します。
   感度低下chがある場合は「ch番号の向き設定」が表示されるので、
   ch1が左端・右端のどちらかを選択します。
6. 必要に応じてS/Nや外観画像、自由記述を記録します。
7. SW Verが「V1.03以降」または「V1.10以降」の場合、
   自由記述欄の下に「プローブ自己診断機能の結果表示」が表示されます。
   プローブ自己診断結果のCSVを1つアップロードすると、
   B11セルが0なら「合格」、1なら「不合格」、
   2なら「プローブ接触不良」と表示されます。
   別のCSVをアップロードすると置き換わり、
   CSVを削除すると、自己診断結果・グラフ・総合判定は非表示になります。
8. プローブ点検結果とプローブ自己診断機能の結果のいずれかが
   不合格の場合は「総合判定結果：不合格」、
   両方とも合格の場合は「総合判定結果：合格」と表示されます。
9. 「この画面をPDFとして保存する」から印刷画面を開き、
   送信先を「PDFに保存」に変更します。
   複数本の場合は、サイドバーの「全プローブのレポートをまとめてPDF出力」で、
   全タブのレポートを1つのPDFにまとめて出力できます
   (プローブごとに改ページされます)。
"""
    )


# =========================================================
# セッション初期化・メイン処理
# =========================================================
# 共通選択(装置本体・SW Ver)
for key, initial_value in {
    "common_device_name": "",
    "common_sw_version": "",
}.items():
    if key not in st.session_state:
        st.session_state[key] = initial_value

# プローブごとの状態
for i in range(max(PROBE_COUNT_OPTIONS)):
    key = f"probename_{i}"
    if key not in st.session_state:
        st.session_state[key] = ""

st.sidebar.title("📋 ページ選択")

selected_page = st.sidebar.radio(
    "ページ",
    options=PAGE_OPTIONS,
    key="page_selector",
    label_visibility="collapsed",
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 📁 超音波画像アップロード")

# 印刷用CSSをページ全体に適用する
inject_print_css()

st.title("🔊 超音波プローブ点検アプリ")

if selected_page == PAGE_HOME:
    render_home()
elif selected_page == PAGE_ABOUT:
    render_about_page()
