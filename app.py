import streamlit as st

from chatgpt_exporter import ChatExportError, export_shared_chat


st.set_page_config(
    page_title="ChatGPT → Word",
    page_icon="💬",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
        .stApp {
            background:
                radial-gradient(circle at 15% 0%, rgba(99,102,241,.15), transparent 28rem),
                radial-gradient(circle at 90% 12%, rgba(16,185,129,.12), transparent 24rem),
                #0b1020;
            color: #f8fafc;
        }

        .block-container {
            max-width: 820px;
            padding-top: 5rem;
            padding-bottom: 4rem;
        }

        .hero {
            text-align: center;
            margin-bottom: 2rem;
        }

        .hero-badge {
            display: inline-block;
            padding: .38rem .75rem;
            border: 1px solid rgba(255,255,255,.12);
            border-radius: 999px;
            background: rgba(255,255,255,.05);
            color: #cbd5e1;
            font-size: .86rem;
            margin-bottom: 1rem;
        }

        .hero h1 {
            margin: 0;
            font-size: clamp(2.35rem, 7vw, 4.8rem);
            letter-spacing: -.055em;
            line-height: .98;
            color: #f8fafc;
        }

        .hero p {
            max-width: 620px;
            margin: 1.15rem auto 0 auto;
            color: #94a3b8;
            font-size: 1.05rem;
            line-height: 1.65;
        }

        .glass-card {
            border: 1px solid rgba(255,255,255,.10);
            background: rgba(15,23,42,.66);
            backdrop-filter: blur(14px);
            border-radius: 24px;
            padding: 1.25rem 1.25rem .75rem 1.25rem;
            box-shadow: 0 24px 80px rgba(0,0,0,.28);
        }

        div[data-testid="stTextInput"] input {
            background: rgba(2,6,23,.72);
            color: #f8fafc;
            border-radius: 14px;
            border: 1px solid rgba(148,163,184,.22);
            min-height: 3.3rem;
        }

        div[data-testid="stTextInput"] input:focus {
            border-color: rgba(99,102,241,.85);
            box-shadow: 0 0 0 1px rgba(99,102,241,.55);
        }

        div.stButton > button,
        div.stDownloadButton > button {
            min-height: 3.2rem;
            border-radius: 14px;
            font-weight: 700;
            border: 1px solid rgba(255,255,255,.12);
        }

        div.stButton > button[kind="primary"] {
            background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 52%, #10b981 130%);
            color: white;
            border: 0;
        }

        .result-card {
            margin-top: 1rem;
            padding: 1rem 1.1rem;
            border-radius: 16px;
            background: rgba(16,185,129,.08);
            border: 1px solid rgba(16,185,129,.22);
        }

        .result-title {
            font-weight: 700;
            color: #ecfdf5;
            margin-bottom: .25rem;
        }

        .result-meta {
            color: #a7f3d0;
            font-size: .92rem;
        }

        .footer-note {
            text-align: center;
            color: #64748b;
            font-size: .82rem;
            margin-top: 1.5rem;
        }

        #MainMenu, footer, header {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)


def clear_export() -> None:
    st.session_state.pop("export_result", None)


st.markdown(
    """
    <div class="hero">
        <div class="hero-badge">SHARED CHAT EXPORTER</div>
        <h1>ChatGPT → Word</h1>
        <p>
            Вставьте публичную share-ссылку на разговор ChatGPT.
            Приложение соберёт сообщения и подготовит аккуратный DOCX для скачивания.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="glass-card">', unsafe_allow_html=True)

shared_link = st.text_input(
    "Shared-ссылка",
    placeholder="https://chatgpt.com/share/…",
    label_visibility="visible",
    on_change=clear_export,
)

convert_clicked = st.button(
    "Создать Word-документ",
    type="primary",
    use_container_width=True,
)

if convert_clicked:
    if not shared_link.strip():
        st.warning("Сначала вставьте shared-ссылку на чат.")
    else:
        try:
            with st.spinner("Загружаю чат и собираю документ…"):
                title, filename, message_count, docx_bytes = export_shared_chat(shared_link)

            st.session_state["export_result"] = {
                "title": title,
                "filename": filename,
                "message_count": message_count,
                "docx_bytes": docx_bytes,
            }
        except ChatExportError as exc:
            st.session_state.pop("export_result", None)
            st.error(str(exc))
        except Exception:
            st.session_state.pop("export_result", None)
            st.error("Произошла непредвиденная ошибка при создании документа.")

result = st.session_state.get("export_result")
if result:
    st.markdown(
        f"""
        <div class="result-card">
            <div class="result-title">Документ готов</div>
            <div class="result-meta">
                {result['title']} · сообщений: {result['message_count']}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.download_button(
        "Скачать .docx",
        data=result["docx_bytes"],
        file_name=result["filename"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )

st.markdown('</div>', unsafe_allow_html=True)

st.markdown(
    """
    <div class="footer-note">
        Работает с публичными ChatGPT share-ссылками. Сам чат никуда дополнительно не сохраняется.
    </div>
    """,
    unsafe_allow_html=True,
)
