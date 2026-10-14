import re
import traceback
from urllib.parse import quote

import requests
import streamlit as st

from generator import build_workbook

st.set_page_config(
    page_title="Verbo Magico",
    page_icon="🇪🇸",
    layout="centered",
)

st.markdown(
    """
    <style>
    .block-container {max-width: 900px; padding-top: 2rem;}
    .hero {
        padding: 1.5rem;
        border-radius: 14px;
        background: linear-gradient(135deg, #17365D, #2F75B5);
        color: white;
        margin-bottom: 1.2rem;
    }
    .hero h1 {margin: 0; font-size: 2rem;}
    .hero p {margin: .45rem 0 0; opacity: .92;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
      <h1>🇪🇸 Verbo Magico</h1>
      <p>Generate a polished Spanish verb workbook from an infinitive.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(ttl=86400, show_spinner=False)
def fetch_conjugation(verb: str) -> dict:
    url = f"https://verbe.cc/verbecc/conjugate/es/{quote(verb, safe='')}"
    response = requests.get(
        url,
        timeout=30,
        headers={
            "Accept": "application/json",
            "User-Agent": "Verbo-Magico-Streamlit/4.0",
        },
    )
    response.raise_for_status()
    return response.json()


verb = st.text_input(
    "Spanish infinitive",
    value="tener",
    placeholder="tener, hablar, ser, ir...",
).strip().lower()

valid = bool(re.fullmatch(r"[a-záéíóúüñ]+", verb)) and verb.endswith(
    ("ar", "er", "ir")
)

st.caption(
    "Includes European Spanish vosotros forms, core tenses, subjunctive, commands and all returned forms."
)

if verb and not valid:
    st.warning("Enter a Spanish infinitive ending in -ar, -er or -ir.")

if st.button(
    "Generate workbook",
    type="primary",
    disabled=not valid,
    use_container_width=True,
):
    try:
        with st.spinner(f"Conjugating {verb} and building the workbook..."):
            payload = fetch_conjugation(verb)
            workbook_bytes = build_workbook(verb, payload)

        if not isinstance(workbook_bytes, bytes) or not workbook_bytes:
            raise TypeError("Workbook generator did not return Excel data.")

        st.success(f"{verb.capitalize()} is ready.")
        st.download_button(
            "📥 Download Excel workbook",
            data=workbook_bytes,
            file_name=f"{verb}_complete_conjugation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except requests.RequestException as exc:
        st.error("The conjugation service could not be reached. Please try again.")
        with st.expander("Technical details"):
            st.code(repr(exc), language="text")

    except Exception:
        st.error("The workbook could not be generated.")
        with st.expander("Technical details"):
            st.code(traceback.format_exc(), language="text")

st.divider()
st.caption("Version 4 • Streamlit Cloud • No Windows installation or macros")
