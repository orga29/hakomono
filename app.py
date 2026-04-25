from pathlib import Path

import streamlit as st

from logic import (
    DEFAULT_SOURCE_SHEET_URL,
    DEFAULT_TEMPLATE_PATH,
    load_source_data,
    process_data,
    suggested_output_filename,
    write_to_template,
)


st.set_page_config(page_title="はこもの集計 2026.04", layout="wide")

st.markdown(
    """
    <style>
    div.stButton > button,
    div.stDownloadButton > button {
        font-weight: 700;
        font-size: 1.08rem;
        padding: 0.7rem 1.3rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("【はこもの集計】2026.04")
st.markdown("元データの Google スプレッドシートを読み込み、ローカルのテンプレート `.xlsx` に書き込んだ完成ファイルをダウンロードできます。")

source_sheet_url = st.text_input(
    "元データ Google スプレッドシート URL",
    value=DEFAULT_SOURCE_SHEET_URL,
)
template_path = Path(DEFAULT_TEMPLATE_PATH)
st.caption(f"テンプレート: {template_path.name}")
st.caption("ダウンロード名は `hakomono-mmdd.xlsx` 形式です。重複時の `(1)` などの付与はブラウザ側の保存動作に従います。")

if st.button("集計開始"):
    try:
        with st.spinner("Processing..."):
            filtered_df, col_mapping = load_source_data(source_sheet_url)
            st.success(
                f"ソースデータを読み込みました。該当する箱ものが{len(filtered_df)}件見つかりました。"
            )

            (
                df_koda,
                df_yamato,
                koda_headers,
                yamato_headers,
                yamato_delivery_types,
            ) = process_data(filtered_df, col_mapping)

            output_buffer = write_to_template(
                template_path,
                df_koda,
                df_yamato,
                koda_headers,
                yamato_headers,
                yamato_delivery_types,
            )

            st.success("集計完了")
            st.download_button(
                label="ファイルをダウンロード",
                data=output_buffer,
                file_name=suggested_output_filename(),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as exc:
        st.error(f"An error occurred: {exc}")
        st.exception(exc)
