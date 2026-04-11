from __future__ import annotations

import hashlib
import io
import json
import os
import uuid
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import matplotlib.pyplot as plt
import streamlit as st

from financial_ratio_core import (
    ALL_METRICS,
    METRIC_GROUPS,
    analyse_requested_tables_with_openai_compatible,
    analyse_requested_tables_with_lmstudio,
    build_analysis_markdown_bundle,
    build_financial_ratio_data,
    build_table_package,
    export_metric_chart_to_svg,
    export_selected_metric_charts_to_svg,
    export_selected_presentable_tables_to_excel,
    get_metric_label,
    get_metric_reference_rows,
    get_pgpass_path,
    has_pgpass_entry,
    presentable_table_to_excel_bytes,
)

st.set_page_config(
    page_title="Financial Ratio Research Studio",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_METRICS = [
    "MARKET_TO_BOOK_RATIO",
    "ROA",
    "ROE",
    "PROFIT_MARGIN",
    "CURRENT_RATIO",
]
DEFAULT_ANALYSIS_STYLE = "entry-level financial analyst"

RUNTIME_MODES = {"local", "cloud"}
DEFAULT_QWEN_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
DEFAULT_QWEN_MODEL = "qwen3.6-plus"
RATE_LIMIT_FILE = Path(".runtime") / "shared_online_llm_limits.json"
RATE_LIMIT_TIMEZONE = ZoneInfo("Asia/Shanghai")


def inject_styles() -> None:
    st.markdown(
        """
<style>
:root {
    --navy: #192638;
    --ink: #122033;
    --gold: #a78649;
    --paper: rgba(252, 249, 242, 0.9);
    --line: rgba(25, 38, 56, 0.12);
}
.stApp {
    background:
        radial-gradient(circle at top left, rgba(255,255,255,0.98), rgba(245,239,228,0.96) 50%, rgba(234,227,214,0.94) 100%),
        linear-gradient(135deg, rgba(25,38,56,0.06), rgba(167,134,73,0.05));
    color: var(--ink);
}
.block-container {
    max-width: 1320px;
    padding-top: 3.4rem;
    padding-bottom: 3rem;
}
h1, h2, h3, h4 {
    color: var(--navy);
    font-family: Georgia, "Times New Roman", serif;
}
.hero-card {
    background: linear-gradient(135deg, rgba(255,255,255,0.88), rgba(247,241,230,0.96));
    border: 1px solid var(--line);
    border-radius: 28px;
    padding: 2.35rem 2.4rem 2.5rem 2.4rem;
    box-shadow: 0 22px 45px rgba(25, 38, 56, 0.08);
    margin-top: 0.45rem;
    margin-bottom: 1.6rem;
}
.hero-kicker {
    color: var(--gold);
    font-size: 0.82rem;
    font-weight: 700;
    letter-spacing: 0.18em;
    text-transform: uppercase;
}
.hero-title {
    color: var(--navy);
    font-family: Georgia, "Times New Roman", serif;
    font-size: 2.5rem;
    line-height: 1.1;
    margin: 0.45rem 0 0.6rem 0;
}
.hero-copy {
    max-width: 850px;
    color: #2d3b4f;
    font-size: 1rem;
    line-height: 1.75;
}
.metric-card {
    background: rgba(255,255,255,0.72);
    border: 1px solid var(--line);
    border-radius: 20px;
    padding: 1rem 1.1rem;
    box-shadow: 0 10px 24px rgba(25, 38, 56, 0.05);
}
.metric-card-label {
    display: block;
    color: #586579;
    font-size: 0.78rem;
    letter-spacing: 0.12em;
    text-transform: uppercase;
}
.metric-card-value {
    display: block;
    color: var(--navy);
    font-family: Georgia, "Times New Roman", serif;
    font-size: 1.5rem;
    margin-top: 0.4rem;
}
.hint-card {
    background: rgba(255,255,255,0.6);
    border-left: 4px solid var(--gold);
    border-radius: 16px;
    padding: 0.9rem 1rem;
    margin-top: 0.7rem;
    margin-bottom: 1.4rem;
    color: #334256;
}
.reference-card {
    background: rgba(255,255,255,0.72);
    border: 1px solid var(--line);
    border-radius: 18px;
    padding: 1rem 1.05rem;
    margin-bottom: 0.9rem;
    box-shadow: 0 10px 22px rgba(25, 38, 56, 0.05);
}
.reference-card-group {
    color: var(--gold);
    font-size: 0.74rem;
    font-weight: 700;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    margin-bottom: 0.45rem;
}
.reference-card-title {
    color: var(--navy);
    font-family: Georgia, "Times New Roman", serif;
    font-size: 1.08rem;
    font-weight: 700;
    margin-bottom: 0.55rem;
}
.reference-card-code {
    color: #607086;
    font-family: "Courier New", monospace;
    font-size: 0.82rem;
    font-weight: 600;
}
.reference-card-row,
.reference-card-note {
    color: #314154;
    font-size: 0.95rem;
    line-height: 1.6;
}
.reference-card-note {
    margin-top: 0.5rem;
    color: #5b6475;
}
div[data-baseweb="input"] > div,
div[data-baseweb="select"] > div,
div[data-testid="stTextArea"] textarea {
    background: rgba(255,255,255,0.88);
}
.stButton > button,
.stDownloadButton > button {
    border: none;
    border-radius: 999px;
    padding: 0.7rem 1.2rem;
    background: linear-gradient(135deg, #192638, #304a63);
    color: white;
    font-weight: 600;
    box-shadow: 0 12px 24px rgba(25, 38, 56, 0.16);
}
.stButton > button:disabled {
    background: linear-gradient(135deg, rgba(25, 38, 56, 0.28), rgba(48, 74, 99, 0.24));
    color: rgba(255, 255, 255, 0.88);
    box-shadow: none;
    cursor: not-allowed;
    opacity: 1;
}
.stTabs [data-baseweb="tab-list"] {
    gap: 0.4rem;
    padding: 0.35rem 0 0.32rem 0;
    overflow-x: auto;
    overflow-y: visible;
}
.stTabs [data-baseweb="tab"] {
    background: rgba(255,255,255,0.62);
    border-radius: 999px;
    padding: 0.72rem 1.35rem;
    min-height: 3.2rem;
    display: flex;
    align-items: center;
}
.stTabs [aria-selected="true"] {
    background: #192638;
    color: white;
}
.stTabs [data-baseweb="tab-border"] {
    margin-top: 0.02rem;
}
.stTabs {
    overflow: visible;
}
</style>
        """,
        unsafe_allow_html=True,
    )


def render_hero() -> None:
    st.markdown(
        """
<div class="hero-card">
    <div class="hero-kicker">WRDS  |  Streamlit  |  LM Studio  |  OpenAI-Compatible APIs</div>
    <div class="hero-title">Financial Ratio Research Studio</div>
    <div class="hero-copy">
        Retrieve Compustat data from WRDS, build publication-style financial tables,
        visualise multi-year trends, and ask either a local model or an online OpenAI-compatible model
        for a concise academic summary. The interface mirrors your notebook workflow while supporting both
        local and cloud deployment modes.
    </div>
</div>
        """,
        unsafe_allow_html=True,
    )


def render_metric_reference() -> None:
    metric_rows = get_metric_reference_rows()

    with st.expander("Metric formulas and meanings", expanded=False):
        st.caption(
            "Search by metric name, code, formula, or meaning if you want a quick reminder before selecting variables."
        )

        filter_col, group_col = st.columns([1.8, 1], gap="large")
        search_text = filter_col.text_input(
            "Search metrics",
            placeholder="ROA, debt, liquidity, turnover, market...",
            key="metric_reference_search",
        ).strip()
        selected_group = group_col.selectbox(
            "Filter by group",
            options=["All groups"] + list(METRIC_GROUPS.keys()),
            index=0,
            key="metric_reference_group",
        )

        filtered_rows = metric_rows
        if selected_group != "All groups":
            filtered_rows = [row for row in filtered_rows if row["group"] == selected_group]

        if search_text:
            query = search_text.lower()
            filtered_rows = [
                row
                for row in filtered_rows
                if query in row["code"].lower()
                or query in row["metric"].lower()
                or query in row["formula"].lower()
                or query in row["meaning"].lower()
                or query in row["notes"].lower()
                or query in row["group"].lower()
            ]

        st.caption(f"{len(filtered_rows)} metric reference item(s) shown.")

        if not filtered_rows:
            st.info("No metrics matched the current search.")
            return

        for group_name in ["All groups"] + list(METRIC_GROUPS.keys()):
            if selected_group != "All groups" and group_name != selected_group:
                continue
            if group_name == "All groups":
                continue

            group_rows = [row for row in filtered_rows if row["group"] == group_name]
            if not group_rows:
                continue

            st.markdown(f"#### {group_name}")
            columns = st.columns(2, gap="large")
            for index, row in enumerate(group_rows):
                with columns[index % 2]:
                    notes_html = (
                        f'<div class="reference-card-note"><strong>Fallback or note:</strong> {row["notes"]}</div>'
                        if row["notes"]
                        else ""
                    )
                    st.markdown(
                        f"""
<div class="reference-card">
    <div class="reference-card-group">{row["group"]}</div>
    <div class="reference-card-title">{row["metric"]} <span class="reference-card-code">({row["code"]})</span></div>
    <div class="reference-card-row"><strong>Formula / basis:</strong> {row["formula"]}</div>
    <div class="reference-card-row"><strong>Meaning:</strong> {row["meaning"]}</div>
    {notes_html}
</div>
                        """,
                        unsafe_allow_html=True,
                    )


def unique_in_order(values: list[str]) -> list[str]:
    return list(dict.fromkeys(values))


def build_metric_defaults(group_metrics: list[str]) -> list[str]:
    return []


def handle_online_provider_preset_change() -> None:
    preset = st.session_state.get("online_provider_preset", "")
    online_defaults = get_online_llm_defaults()

    if preset == "Custom OpenAI-compatible API":
        st.session_state["online_api_base_url"] = ""
        st.session_state["online_api_model"] = ""
        st.session_state["online_api_key"] = ""
        st.session_state["online_enable_thinking"] = False
    elif preset == "Alibaba Bailian / Qwen 3.6 Plus":
        st.session_state["online_api_base_url"] = online_defaults["base_url"]
        st.session_state["online_api_model"] = online_defaults["model"]
        st.session_state["online_api_key"] = ""
        st.session_state["online_enable_thinking"] = online_defaults["enable_thinking"]
        st.session_state["online_use_system_proxy"] = online_defaults["use_system_proxy"]


def render_stat_card(label: str, value: str) -> None:
    st.markdown(
        f"""
<div class="metric-card">
    <span class="metric-card-label">{label}</span>
    <span class="metric-card-value">{value}</span>
</div>
        """,
        unsafe_allow_html=True,
    )


def prepare_widget_state() -> None:
    if st.session_state.pop("_clear_wrds_password", False):
        st.session_state.pop("wrds_password", None)


def render_run_notice() -> None:
    notice = st.session_state.pop("run_notice", None)
    if not notice:
        return

    level = notice.get("level")
    message = notice.get("message", "")

    if level == "success":
        st.success(message)
    elif level == "warning":
        st.warning(message)
    elif level == "error":
        st.error(message)
    else:
        st.info(message)


def get_secret_value(name: str, default=None):
    env_value = os.getenv(name)
    if env_value is not None:
        return env_value
    try:
        return st.secrets[name]
    except Exception:
        return default


def is_probably_streamlit_cloud_host() -> bool:
    host = ""
    try:
        host = str(st.context.headers.get("host", "") or "").lower()
    except Exception:
        host = ""

    return host.endswith(".streamlit.app") or "share.streamlit.io" in host


def get_runtime_mode() -> str:
    configured = str(get_secret_value("APP_RUNTIME_MODE", "local")).strip().lower()
    if configured in RUNTIME_MODES:
        return configured
    if is_probably_streamlit_cloud_host():
        return "cloud"
    return "local"


def is_cloud_mode() -> bool:
    return get_runtime_mode() == "cloud"


def get_online_llm_defaults() -> dict:
    return {
        "provider_label": str(
            get_secret_value("DEFAULT_ONLINE_LLM_PROVIDER_LABEL", "Alibaba Bailian / Qwen 3.6 Plus")
        ),
        "base_url": str(get_secret_value("DEFAULT_ONLINE_LLM_BASE_URL", DEFAULT_QWEN_BASE_URL)),
        "model": str(get_secret_value("DEFAULT_ONLINE_LLM_MODEL", DEFAULT_QWEN_MODEL)),
        "api_key": str(get_secret_value("DEFAULT_ONLINE_LLM_API_KEY", "")),
        "enable_thinking": str(get_secret_value("DEFAULT_ONLINE_LLM_ENABLE_THINKING", "true")).lower()
        in {"1", "true", "yes", "on"},
        "use_system_proxy": str(get_secret_value("DEFAULT_ONLINE_LLM_USE_SYSTEM_PROXY", "false")).lower()
        in {"1", "true", "yes", "on"},
        "daily_limit": int(get_secret_value("SHARED_ONLINE_LLM_DAILY_LIMIT", 3)),
    }


def get_request_fingerprint() -> str:
    ip_address = getattr(st.context, "ip_address", "") or ""
    user_agent = ""
    try:
        user_agent = st.context.headers.get("user-agent", "")
    except Exception:
        user_agent = ""

    raw_identity = f"{ip_address}|{user_agent}".strip("|")
    if not raw_identity:
        raw_identity = st.session_state.setdefault("_anonymous_request_id", uuid.uuid4().hex)

    return hashlib.sha256(raw_identity.encode("utf-8")).hexdigest()


def load_rate_limit_data() -> dict:
    if not RATE_LIMIT_FILE.exists():
        return {}
    try:
        return json.loads(RATE_LIMIT_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_rate_limit_data(data: dict) -> None:
    RATE_LIMIT_FILE.parent.mkdir(parents=True, exist_ok=True)
    RATE_LIMIT_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")


def get_shared_online_quota_status(limit: int) -> tuple[int, int]:
    today = datetime.now(RATE_LIMIT_TIMEZONE).date().isoformat()
    fingerprint = get_request_fingerprint()
    data = load_rate_limit_data()
    used = int(data.get(today, {}).get(fingerprint, 0))
    return used, max(limit - used, 0)


def consume_shared_online_quota(limit: int) -> tuple[bool, int, int]:
    today = datetime.now(RATE_LIMIT_TIMEZONE).date().isoformat()
    fingerprint = get_request_fingerprint()
    data = load_rate_limit_data()

    # Keep only recent dates so the file stays small.
    recent_dates = sorted(data.keys())[-6:]
    data = {date_key: data[date_key] for date_key in recent_dates if date_key in data}

    day_bucket = data.setdefault(today, {})
    used = int(day_bucket.get(fingerprint, 0))
    if used >= limit:
        return False, used, limit

    day_bucket[fingerprint] = used + 1
    try:
        save_rate_limit_data(data)
    except Exception:
        session_key = "_shared_online_quota_used"
        st.session_state[session_key] = int(st.session_state.get(session_key, 0)) + 1
        used = int(st.session_state.get(session_key, 0))
        if used > limit:
            return False, limit, limit
        return True, used, limit

    return True, used + 1, limit


def collect_exception_messages(exc: Exception) -> list[str]:
    messages: list[str] = []
    seen: set[int] = set()
    current: Exception | None = exc
    while current is not None and id(current) not in seen:
        seen.add(id(current))
        message = str(current).strip()
        if message and message not in messages:
            messages.append(message)
        next_exc = getattr(current, "__cause__", None) or getattr(current, "__context__", None)
        current = next_exc if isinstance(next_exc, Exception) else None
    return messages


def get_online_provider_name(form_values: dict) -> str:
    if form_values.get("online_provider_preset") == "Alibaba Bailian / Qwen 3.6 Plus":
        return "Alibaba Bailian / Qwen 3.6 Plus"
    return "Custom OpenAI-compatible API"


def get_online_api_style(form_values: dict) -> str:
    if form_values.get("online_provider_preset") == "Alibaba Bailian / Qwen 3.6 Plus":
        return "responses"
    return "chat_completions"


def build_online_llm_error_message(
    exc: Exception,
    *,
    provider_name: str = "",
    base_url: str = "",
    model_id: str = "",
) -> str:
    messages = collect_exception_messages(exc)
    message = messages[0] if messages else exc.__class__.__name__
    lowered = " ".join(messages).lower()
    details = [message]

    if len(messages) > 1:
        root_message = messages[-1]
        if root_message and root_message != message:
            details.append(f"Underlying error: {root_message}")

    target_parts = []
    if provider_name:
        target_parts.append(provider_name)
    if model_id:
        target_parts.append(f"model `{model_id}`")
    if base_url:
        target_parts.append(f"base URL `{base_url}`")
    if target_parts:
        details.append("Request target: " + " | ".join(target_parts))

    quota_markers = ["quota", "balance", "insufficient", "credit", "429", "402", "rate limit"]
    if any(marker in lowered for marker in quota_markers):
        details.append(
            "The shared online API quota may be exhausted. Please try LM Studio, "
            "switch to another online model, or enter your own API key and custom base URL."
        )
    elif any(
        marker in lowered
        for marker in ["connection error", "connect", "timed out", "timeout", "ssl", "certificate", "dns"]
    ):
        details.append(
            "The app reached the online-provider step, but the request could not be completed. "
            "If this keeps happening, try your own API key, confirm the base URL is reachable from this computer, "
            "or switch to LM Studio."
        )

    return "\n\n".join(details)


def build_form() -> dict:
    metric_choices = {}
    runtime_mode = get_runtime_mode()
    cloud_mode = runtime_mode == "cloud"
    online_defaults = get_online_llm_defaults()

    st.subheader("Research Setup")
    left_col, right_col = st.columns([1, 1.15], gap="large")

    with left_col:
        st.markdown("#### WRDS Access")
        username = st.text_input("WRDS username", key="wrds_username")
        username_clean = username.strip()
        pgpass_ready = (not cloud_mode) and bool(username_clean) and has_pgpass_entry(username_clean)
        password_placeholder = (
            "Optional: if pgpass already exists."
            if not cloud_mode
            else "Enter password for this session."
        )
        password_help = (
            "If a valid pgpass entry already exists for this WRDS username, "
            "you can leave this box empty. Enter the password only when pgpass "
            "has not been created yet or you want to create/update it now."
            if not cloud_mode
            else "In cloud mode, pgpass creation is disabled. Enter the WRDS password for this session."
        )
        password = st.text_input(
            "WRDS password",
            type="password",
            key="wrds_password",
            placeholder=password_placeholder,
            help=password_help,
        )
        if not cloud_mode:
            st.caption("Password is optional when a valid `pgpass.conf` entry already exists for this username.")

        if cloud_mode:
            create_pgpass = False
        else:
            if pgpass_ready:
                st.caption(f"A pgpass entry was found for `{username_clean}`. The password can be left blank.")
            elif username_clean:
                st.caption(
                    "No pgpass entry was detected for this username, so enter the password if you want WRDS retrieval to run."
                )
            else:
                st.caption("Tip: once pgpass.conf is created, you can leave the password blank in future runs.")
            create_pgpass = st.checkbox(
                "Create or update pgpass.conf for future WRDS logins",
                key="create_pgpass",
            )
            st.caption(f"Current pgpass location: `{get_pgpass_path()}`")

        st.markdown("#### AI Summary")
        ai_provider_options = (
            ["Disabled", "Online OpenAI-compatible API", "LM Studio (local)"]
            if not cloud_mode
            else ["Disabled", "Online OpenAI-compatible API"]
        )
        default_ai_provider = "Online OpenAI-compatible API"
        ai_provider = st.selectbox(
            "Summary provider",
            options=ai_provider_options,
            index=ai_provider_options.index(default_ai_provider),
            key="ai_provider",
            help=(
                "Use LM Studio for a fully local summary, or use an online OpenAI-compatible API such as Alibaba Bailian."
                if not cloud_mode
                else "Community Cloud cannot access your own localhost, so online API mode is used for hosted deployment."
            ),
        )

        if ai_provider == "LM Studio (local)":
            st.caption("LM Studio mode sends the markdown tables to your local OpenAI-compatible server.")
            lmstudio_url = st.text_input(
                "LM Studio API URL",
                value="http://localhost:1234/v1",
                key="lmstudio_url",
            )
            preferred_model = st.text_input(
                "Preferred model id",
                value="gemma-4-e4b-it",
                key="preferred_model",
            )
            online_api_base_url = ""
            online_api_model = ""
            online_api_key = ""
            online_enable_thinking = False
            online_use_system_proxy = False
            online_provider_preset = ""
        elif ai_provider == "Online OpenAI-compatible API":
            preset_options = ["Alibaba Bailian / Qwen 3.6 Plus", "Custom OpenAI-compatible API"]
            default_preset = preset_options[0]
            online_provider_preset = st.selectbox(
                "Online provider preset",
                options=preset_options,
                index=preset_options.index(default_preset),
                key="online_provider_preset",
                on_change=handle_online_provider_preset_change,
            )
            if online_provider_preset == "Alibaba Bailian / Qwen 3.6 Plus":
                st.caption(
                    "Recommended online default. Thinking mode is supported. If the shared quota is exhausted, "
                    "switch to LM Studio or enter your own API key and custom model."
                )
            else:
                st.caption(
                    "Use any OpenAI-compatible endpoint. The app will keep the same financial-analysis prompt so answers stay consistent."
                )

            online_api_base_url = st.text_input(
                "Online API base URL",
                value=online_defaults["base_url"],
                key="online_api_base_url",
            )
            online_api_model = st.text_input(
                "Online model id",
                value=online_defaults["model"],
                key="online_api_model",
            )
            online_api_key = st.text_input(
                "Online API key",
                type="password",
                key="online_api_key",
                placeholder="Optional: blank uses shared key.",
                help=(
                    "If you leave this blank, the app will try the shared server-side API key configured by the app owner. "
                    "You can also enter your own API key, model, and base URL."
                ),
            )
            if online_provider_preset == "Custom OpenAI-compatible API":
                st.caption("Custom preset clears the base URL, model id, and key so users can enter their own values from scratch.")
            elif bool(online_defaults["api_key"]):
                st.caption("Leave the key blank to use the app's shared online key, or enter your own key for this provider.")
            else:
                st.caption("Enter your own online API key for this provider.")
            online_enable_thinking = st.checkbox(
                "Enable thinking mode",
                value=online_defaults["enable_thinking"],
                key="online_enable_thinking",
                help="For providers such as Qwen on Alibaba Bailian, this sends the non-standard enable_thinking flag through the OpenAI-compatible request.",
            )
            with st.expander("Advanced online connection settings", expanded=False):
                online_use_system_proxy = st.checkbox(
                    "Use system proxy / environment network settings",
                    value=online_defaults["use_system_proxy"],
                    key="online_use_system_proxy",
                    help=(
                        "Turn this on only if your network requires proxy environment variables such as "
                        "HTTP_PROXY or HTTPS_PROXY. Most users should leave it off."
                    ),
                )
                st.caption(
                    "Direct connection is recommended by default. Enable system proxy only when your "
                    "campus, company, or VPN environment requires it."
                )
            shared_key_available = bool(online_defaults["api_key"])
            if shared_key_available:
                used, remaining = get_shared_online_quota_status(online_defaults["daily_limit"])
                st.caption(
                    f"Shared online key status: {remaining} of {online_defaults['daily_limit']} summary request(s) remaining today on this connection."
                )
            else:
                st.caption("No shared online API key is configured, so users should enter their own key for online summaries.")

            lmstudio_url = ""
            preferred_model = ""
        else:
            lmstudio_url = ""
            preferred_model = ""
            online_api_base_url = ""
            online_api_model = ""
            online_api_key = ""
            online_enable_thinking = False
            online_use_system_proxy = False
            online_provider_preset = ""

    with right_col:
        st.markdown("#### Search Scope")
        tickers_text = st.text_area(
            "Ticker codes",
            height=110,
            placeholder="AAPL, NVDA, MSFT",
            key="tickers_text",
        )
        sic_code = st.text_input(
            "Industry SIC code (optional benchmark)",
            key="sic_code",
        )
        year_col1, year_col2 = st.columns(2)
        start_year = year_col1.number_input(
            "Start year",
            min_value=1960,
            max_value=2100,
            value=2019,
            step=1,
            key="start_year",
        )
        end_year = year_col2.number_input(
            "End year",
            min_value=1960,
            max_value=2100,
            value=2024,
            step=1,
            key="end_year",
        )
        cost_of_capital = st.number_input(
            "Cost of capital for EVA",
            min_value=0.00,
            max_value=1.00,
            value=0.10,
            step=0.01,
            format="%.2f",
            key="cost_of_capital",
        )
        temperature = st.slider(
            "Summary temperature",
            min_value=0.0,
            max_value=1.0,
            value=0.2,
            step=0.1,
            key="temperature",
        )
        max_tokens = st.slider(
            "Maximum summary tokens",
            min_value=400,
            max_value=2200,
            value=1000,
            step=100,
            key="max_tokens",
        )

        if cloud_mode:
            st.info(
                "Cloud mode is active: path-based saving is disabled, browser downloads remain available, and LM Studio is replaced by online API mode."
            )
        else:
            st.info(
                "Local mode is active: pgpass creation, LM Studio, and save-to-folder exports are available alongside online API support."
            )

    st.markdown("#### Financial Metrics")
    select_all_metrics = st.checkbox("Include every metric", key="select_all_metrics")
    st.caption("Metric selections start blank by default so users can choose only the ratios they actually need.")

    for group_name, group_metrics in METRIC_GROUPS.items():
        group_key = f"group_{group_name.lower().replace(' ', '_').replace('-', '_')}"
        with st.expander(group_name, expanded=group_name in {"Performance Measures", "Profitability Measures"}):
            metric_choices[group_name] = st.multiselect(
                f"{group_name} options",
                options=group_metrics,
                default=build_metric_defaults(group_metrics),
                format_func=get_metric_label,
                key=group_key,
            )

    selected_metric_count = sum(len(metric_choices.get(group_name, [])) for group_name in METRIC_GROUPS)
    if not select_all_metrics and selected_metric_count == 0:
        st.caption("Choose at least one financial metric before running the report.")

    metrics_ready = bool(select_all_metrics or selected_metric_count > 0)
    years_ready = int(start_year) <= int(end_year)
    tickers_ready = bool(tickers_text.strip())
    username_ready = bool(username_clean)
    password_ready = bool(password.strip())

    if cloud_mode:
        wrds_ready = bool(username_ready and password_ready)
    else:
        wrds_ready = bool(username_ready and (password_ready or pgpass_ready))

    if ai_provider == "Online OpenAI-compatible API":
        online_base_ready = bool((online_api_base_url or "").strip())
        online_model_ready = bool((online_api_model or "").strip())
        online_key_ready = bool((online_api_key or "").strip()) or bool(online_defaults["api_key"])
        ai_ready = bool(online_base_ready and online_model_ready and online_key_ready)
    elif ai_provider == "LM Studio (local)":
        ai_ready = bool((lmstudio_url or "").strip() and (preferred_model or "").strip())
    else:
        ai_ready = True

    form_ready = bool(wrds_ready and tickers_ready and metrics_ready and years_ready and ai_ready)

    submitted = st.button(
        "Retrieve data and build report",
        use_container_width=True,
        type="primary",
        disabled=not form_ready,
    )

    return {
        "submitted": submitted,
        "runtime_mode": runtime_mode,
        "username": username,
        "password": password,
        "create_pgpass": create_pgpass,
        "tickers_text": tickers_text,
        "sic_code": sic_code,
        "start_year": int(start_year),
        "end_year": int(end_year),
        "cost_of_capital": float(cost_of_capital),
        "select_all_metrics": select_all_metrics,
        "metric_choices": metric_choices,
        "ai_provider": ai_provider,
        "lmstudio_url": lmstudio_url,
        "preferred_model": preferred_model,
        "online_provider_preset": online_provider_preset,
        "online_api_base_url": online_api_base_url,
        "online_api_model": online_api_model,
        "online_api_key": online_api_key,
        "online_enable_thinking": online_enable_thinking,
        "online_use_system_proxy": online_use_system_proxy,
        "analysis_style": DEFAULT_ANALYSIS_STYLE,
        "temperature": float(temperature),
        "max_tokens": int(max_tokens),
    }


def run_report(form_values: dict) -> None:
    selected_metrics = []
    if form_values["select_all_metrics"]:
        selected_metrics = list(ALL_METRICS)
    else:
        for group_name in METRIC_GROUPS:
            selected_metrics.extend(form_values["metric_choices"].get(group_name, []))
        selected_metrics = unique_in_order(selected_metrics)

    if not selected_metrics:
        st.error("Please choose at least one financial metric before running the report.")
        return

    analysis_result = None
    analysis_error = None
    online_defaults = get_online_llm_defaults()
    using_shared_online_key = False
    online_provider_name = ""
    online_api_style = ""

    try:
        st.info(
            "If WRDS asks for Duo 2FA, finish the authentication first. "
            "WRDS data cannot be loaded until the Duo approval is completed."
        )
        with st.spinner("Retrieving WRDS data, building tables, and preparing charts..."):
            report = build_financial_ratio_data(
                username=form_values["username"],
                password=form_values["password"] or None,
                create_pgpass=form_values["create_pgpass"] if form_values["runtime_mode"] == "local" else False,
                tickers=form_values["tickers_text"],
                start_year=form_values["start_year"],
                end_year=form_values["end_year"],
                metrics=selected_metrics,
                sic_code=form_values["sic_code"] or None,
                cost_of_capital=form_values["cost_of_capital"] if "EVA" in selected_metrics else None,
            )
            report = build_table_package(report)

        for key in ["saved_excel_files", "saved_svg_files", "saved_markdown_path"]:
            st.session_state.pop(key, None)

        if form_values["ai_provider"] != "Disabled" and report["pivot_tables"]:
            if form_values["ai_provider"] == "LM Studio (local)":
                with st.spinner("Asking LM Studio for a compact summary..."):
                    try:
                        analysis_result = analyse_requested_tables_with_lmstudio(
                            report=report,
                            metrics=selected_metrics,
                            preferred_model=form_values["preferred_model"],
                            base_url=form_values["lmstudio_url"],
                            temperature=form_values["temperature"],
                            max_tokens=form_values["max_tokens"],
                            analysis_style=form_values["analysis_style"],
                            save_result=False,
                        )
                    except Exception as exc:
                        analysis_error = str(exc)
            elif form_values["ai_provider"] == "Online OpenAI-compatible API":
                user_online_key = form_values["online_api_key"].strip()
                effective_online_key = user_online_key or online_defaults["api_key"]
                using_shared_online_key = not user_online_key and bool(online_defaults["api_key"])
                online_provider_name = get_online_provider_name(form_values)
                online_api_style = get_online_api_style(form_values)

                if not effective_online_key:
                    analysis_error = (
                        "No online API key is available. Enter your own API key, or configure a shared online API key in the app secrets."
                    )
                else:
                    if using_shared_online_key:
                        used, remaining = get_shared_online_quota_status(online_defaults["daily_limit"])
                        if remaining <= 0:
                            analysis_error = (
                                f"The shared online summary quota is limited to {online_defaults['daily_limit']} request(s) per day on this connection, "
                                "and today's limit has already been used. Please try LM Studio, enter your own API key, "
                                "or switch to another online model."
                            )

                    if analysis_error is None:
                        with st.spinner("Asking the online OpenAI-compatible model for a compact summary..."):
                            try:
                                analysis_result = analyse_requested_tables_with_openai_compatible(
                                    report=report,
                                    metrics=selected_metrics,
                                    model_id=form_values["online_api_model"],
                                    base_url=form_values["online_api_base_url"],
                                    api_key=effective_online_key,
                                    temperature=form_values["temperature"],
                                    max_tokens=form_values["max_tokens"],
                                    analysis_style=form_values["analysis_style"],
                                    save_result=False,
                                    enable_thinking=form_values["online_enable_thinking"],
                                    use_system_proxy=form_values["online_use_system_proxy"],
                                    provider_name=online_provider_name,
                                    api_style=online_api_style,
                                )
                                if using_shared_online_key:
                                    allowed, used_after, limit = consume_shared_online_quota(online_defaults["daily_limit"])
                                    if allowed:
                                        st.session_state["shared_online_quota_last_used"] = used_after
                                    else:
                                        analysis_error = (
                                            f"The shared online summary quota is limited to {limit} request(s) per day on this connection. "
                                            "The summary finished, but no further shared-key requests are available today."
                                        )
                            except Exception as exc:
                                analysis_error = build_online_llm_error_message(
                                    exc,
                                    provider_name=online_provider_name,
                                    base_url=form_values["online_api_base_url"],
                                    model_id=form_values["online_api_model"],
                                )

        st.session_state["app_report"] = report
        st.session_state["selected_metrics"] = selected_metrics
        st.session_state["analysis_result"] = analysis_result
        st.session_state["analysis_error"] = analysis_error
        st.session_state["query_context"] = {
            "tickers": form_values["tickers_text"],
            "sic_code": form_values["sic_code"],
            "start_year": form_values["start_year"],
            "end_year": form_values["end_year"],
            "runtime_mode": form_values["runtime_mode"],
            "ai_provider": form_values["ai_provider"],
            "using_shared_online_key": using_shared_online_key,
            "online_provider_name": online_provider_name,
            "online_api_base_url": form_values["online_api_base_url"],
            "online_api_model": form_values["online_api_model"],
            "online_api_style": online_api_style,
            "online_use_system_proxy": form_values["online_use_system_proxy"],
        }

        if report["pivot_tables"]:
            st.session_state["run_notice"] = {
                "level": "success",
                "message": "The report is ready. Explore the data, tables, charts, and summary below.",
            }
        else:
            st.session_state["run_notice"] = {
                "level": "warning",
                "message": "WRDS returned no records for the current query.",
            }

        st.session_state["_clear_wrds_password"] = True
        st.rerun()
    except Exception as exc:
        st.error(str(exc))


def render_results(report: dict, selected_metrics: list[str], analysis_result: dict | None, analysis_error: str | None) -> None:
    query_context = st.session_state.get("query_context", {})
    cloud_mode = query_context.get("runtime_mode") == "cloud"

    company_count = report["long_table"].loc[report["long_table"]["tic"] != "", "tic"].nunique()
    benchmark_label = "Included" if query_context.get("sic_code") else "None"
    year_range = f"{query_context.get('start_year', '-')}-{query_context.get('end_year', '-')}"
    metric_count = str(len(selected_metrics))

    stat_cols = st.columns(4)
    with stat_cols[0]:
        render_stat_card("Companies", str(company_count))
    with stat_cols[1]:
        render_stat_card("Benchmark SIC", benchmark_label)
    with stat_cols[2]:
        render_stat_card("Year Range", year_range)
    with stat_cols[3]:
        render_stat_card("Metrics", metric_count)

    st.markdown(
        """
<div class="hint-card">
    Start and end values are treated as fiscal years. The industry benchmark is appended only when an SIC code
    is supplied and WRDS returns enough observations for the benchmark calculation.
</div>
        """,
        unsafe_allow_html=True,
    )

    if not report["pivot_tables"]:
        st.info("No metric tables were generated for this query. Try adjusting the tickers, years, or SIC benchmark.")
        raw_tabs = st.tabs(["Company Fundamentals", "Industry Benchmark"])
        with raw_tabs[0]:
            if report["company_raw"].empty:
                st.info("No company fundamentals were returned from WRDS.")
            else:
                st.dataframe(report["company_raw"], use_container_width=True, hide_index=True)
        with raw_tabs[1]:
            if report["industry_raw"].empty:
                st.info("No industry benchmark rows were retrieved for this query.")
            else:
                st.dataframe(report["industry_raw"], use_container_width=True, hide_index=True)
        return

    tables_tab, charts_tab, ai_tab, exports_tab, overview_tab = st.tabs(
        ["Formatted Tables", "Visualisations", "AI Summary", "Exports", "Retrieved Data"]
    )

    with tables_tab:
        table_metric = st.selectbox(
            "Choose a metric table",
            options=selected_metrics,
            format_func=get_metric_label,
            key="table_metric",
        )
        st.markdown(f"### {report['titles'][table_metric]}")
        st.table(report["display_tables"][table_metric])
        st.caption(report["notes"][table_metric])
        st.caption(report["sources"][table_metric])

        excel_bytes = presentable_table_to_excel_bytes(report["presentable_tables"][table_metric])
        markdown_text = report["markdown_tables"][table_metric]
        download_cols = st.columns(2)
        with download_cols[0]:
            st.download_button(
                "Download current table as Excel",
                data=excel_bytes,
                file_name=f"{table_metric.lower()}_table.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with download_cols[1]:
            st.download_button(
                "Download current table as markdown",
                data=markdown_text.encode("utf-8"),
                file_name=f"{table_metric.lower()}_table.md",
                mime="text/markdown",
                use_container_width=True,
            )

        with st.expander("View compact markdown table"):
            st.code(markdown_text, language="markdown")

    with charts_tab:
        chart_metric = st.selectbox(
            "Choose a chart metric",
            options=selected_metrics,
            format_func=get_metric_label,
            key="chart_metric",
        )
        fig, _ = export_metric_chart_to_svg(report=report, metric=chart_metric)
        buffer = io.BytesIO()
        fig.savefig(buffer, format="svg", bbox_inches="tight")
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)

        st.download_button(
            "Download current chart as SVG",
            data=buffer.getvalue(),
            file_name=f"{chart_metric.lower()}_chart.svg",
            mime="image/svg+xml",
            use_container_width=True,
        )

    with ai_tab:
        if analysis_result:
            if analysis_result.get("provider_name"):
                st.caption(
                    f"Summary provider: {analysis_result['provider_name']}"
                    + (" with thinking mode enabled." if analysis_result.get("enable_thinking") else ".")
                )
            st.markdown(analysis_result["analysis"])
            ai_cols = st.columns(2)
            with ai_cols[0]:
                st.download_button(
                    "Download AI summary",
                    data=analysis_result["analysis"].encode("utf-8"),
                    file_name="financial_summary.md",
                    mime="text/markdown",
                    use_container_width=True,
                )
            with ai_cols[1]:
                st.download_button(
                    "Download combined markdown tables",
                    data=analysis_result["combined_tables_text"].encode("utf-8"),
                    file_name="combined_metric_tables.md",
                    mime="text/markdown",
                    use_container_width=True,
                )
            with st.expander("Prompt sent to the model"):
                st.code(analysis_result["prompt"], language="markdown")
        elif analysis_error:
            st.warning(analysis_error)
            if query_context.get("ai_provider") == "Online OpenAI-compatible API":
                provider_bits = []
                if query_context.get("online_provider_name"):
                    provider_bits.append(str(query_context["online_provider_name"]))
                if query_context.get("online_api_model"):
                    provider_bits.append(f"model `{query_context['online_api_model']}`")
                if query_context.get("online_api_base_url"):
                    provider_bits.append(f"base URL `{query_context['online_api_base_url']}`")
                if query_context.get("online_api_style"):
                    provider_bits.append(f"API style `{query_context['online_api_style']}`")
                provider_bits.append(
                    "system proxy `on`" if query_context.get("online_use_system_proxy") else "system proxy `off`"
                )
                if provider_bits:
                    st.caption("Attempted online request: " + " | ".join(provider_bits))
        else:
            st.info("Choose LM Studio or an online OpenAI-compatible API before running the report to generate a summary.")

    with exports_tab:
        if cloud_mode:
            st.info(
                "Cloud mode uses browser downloads rather than server-side save paths. "
                "Use the download buttons in the Tables, Visualisations, and AI Summary tabs."
            )
            st.caption(
                "This avoids saving files into the temporary server filesystem on Streamlit Community Cloud."
            )
        else:
            excel_dir = st.text_input(
                "Folder for Excel exports",
                value=str(Path.cwd() / "exports" / "excel_tables"),
                key="excel_export_dir",
            )
            svg_dir = st.text_input(
                "Folder for SVG exports",
                value=str(Path.cwd() / "exports" / "svg_charts"),
                key="svg_export_dir",
            )
            summary_path = st.text_input(
                "Markdown path for the AI summary",
                value=str(Path.cwd() / "exports" / "financial_summary.md"),
                key="summary_export_path",
            )

            export_cols = st.columns(3)
            with export_cols[0]:
                if st.button("Save Excel tables locally", use_container_width=True):
                    saved_excel = export_selected_presentable_tables_to_excel(
                        report=report,
                        metrics=selected_metrics,
                        output_folder=excel_dir,
                    )
                    st.session_state["saved_excel_files"] = saved_excel
            with export_cols[1]:
                if st.button("Save SVG charts locally", use_container_width=True):
                    saved_svg = export_selected_metric_charts_to_svg(
                        report=report,
                        metrics=selected_metrics,
                        output_folder=svg_dir,
                    )
                    st.session_state["saved_svg_files"] = saved_svg
            with export_cols[2]:
                if st.button("Save markdown package locally", use_container_width=True):
                    export_path = Path(summary_path)
                    export_path.parent.mkdir(parents=True, exist_ok=True)

                    markdown_sections = [build_analysis_markdown_bundle(report, selected_metrics)]
                    if analysis_result:
                        markdown_sections.extend(["", "# AI Summary", "", analysis_result["analysis"]])

                    export_path.write_text("\n".join(markdown_sections), encoding="utf-8")
                    st.session_state["saved_markdown_path"] = str(export_path)

            if st.session_state.get("saved_excel_files"):
                st.caption("Saved Excel files")
                st.code("\n".join(st.session_state["saved_excel_files"].values()))

            if st.session_state.get("saved_svg_files"):
                st.caption("Saved SVG files")
                st.code("\n".join(st.session_state["saved_svg_files"].values()))

            if st.session_state.get("saved_markdown_path"):
                st.caption("Saved markdown file")
                st.code(st.session_state["saved_markdown_path"])

    with overview_tab:
        raw_tabs = st.tabs(["Calculated Long Table", "Company Fundamentals", "Industry Benchmark"])
        with raw_tabs[0]:
            st.dataframe(report["long_table_display"], use_container_width=True, hide_index=True)
        with raw_tabs[1]:
            if report["company_raw"].empty:
                st.info("No company fundamentals were returned from WRDS.")
            else:
                st.dataframe(report["company_raw"], use_container_width=True, hide_index=True)
        with raw_tabs[2]:
            if report["industry_raw"].empty:
                st.info("No industry benchmark rows were retrieved for this query.")
            else:
                st.dataframe(report["industry_raw"], use_container_width=True, hide_index=True)


inject_styles()
render_hero()
render_metric_reference()
prepare_widget_state()
render_run_notice()

with st.sidebar:
    st.markdown("### Workflow")
    st.markdown(
        """
1. Enter your WRDS username and password.
2. Choose ticker codes, fiscal years, and financial metrics.
3. Optionally add an SIC benchmark and choose an AI summary provider.
4. Run the report, then export the tables, charts, or summary.
        """
    )
    st.markdown("### Notes")
    st.caption("The app uses one shared codebase with two runtime modes: local and cloud.")
    if is_cloud_mode():
        st.caption("Cloud mode is active: use browser downloads and online OpenAI-compatible APIs.")
    else:
        st.caption("Local mode is active: pgpass, LM Studio, online APIs, and folder exports are available.")
        st.caption(f"`pgpass.conf` will be created at `{get_pgpass_path()}` when requested.")

form_values = build_form()

if form_values["submitted"]:
    run_report(form_values)

current_report = st.session_state.get("app_report")
if current_report:
    render_results(
        report=current_report,
        selected_metrics=st.session_state.get("selected_metrics", []),
        analysis_result=st.session_state.get("analysis_result"),
        analysis_error=st.session_state.get("analysis_error"),
    )
else:
    st.info("Enter your WRDS account, query settings, and chosen metrics above to build the report.")
