# 対訳ファイル作成アプリ
# 和文と英文の統合報告書PDFから対訳ファイルのエクセルを作成します
# 2024-06-17
# openpyxlライブラリのインストールが必要
# 使い方
# 1. 和文PDFをアップロード
# 2. 英文PDFをアップロード
# 3. 対訳ファイルが作成されるので、ダウンロードボタンを押して保存してください

import pymupdf  # PyMuPDFのライブラリ
import pandas as pd
import streamlit as st
import os

def extract_paragraphs_to_file(doc_ja, doc_en, output_xlsx):

    # PDFファイルを開く
    # doc_ja = pymupdf.open(pdf_path_ja)
    # doc_en = pymupdf.open(pdf_path_en)
		
    df = pd.DataFrame()

	# ページごとにループ
    for page_num, page in enumerate(doc_ja):
        list_ja, list_en = [], []         
    	# get_text("blocks")で段落（ブロック）ごとにテキストを取得
    	# sort=True を設定すると上から下の順序で取得
        blocks_ja = doc_ja[page_num].get_text("blocks", sort=True)
        list_ja.append("\n"+"\n"+"\n"+"P"+str(page_num))
            
        for block_ja in blocks_ja:
        # blockは (x0, y0, x1, y1, text, block_no, block_type) のタプル
        # block[4] にテキストが含まれる
            if block_ja[6] == 0:  # block_type == 0 はテキストブロック
                text_ja = block_ja[4].strip()
                if text_ja:  # 空白でない場合のみ出力
                    list_ja.append("\n")
                    list_ja.append(text_ja)

    	# get_text("blocks")で段落（ブロック）ごとにテキストを取得
    	# sort=True を設定すると上から下の順序で取得
        blocks_en = doc_en[page_num].get_text("blocks", sort=True)
        list_en.append("\n"+"\n"+"\n"+"P"+str(page_num))
            
        for block_en in blocks_en:
        # blockは (x0, y0, x1, y1, text, block_no, block_type) のタプル
        # block[4] にテキストが含まれる
            if block_en[6] == 0:  # block_type == 0 はテキストブロック
                text_en = block_en[4].strip()
                if text_en:  # 空白でない場合のみ出力
                    list_en.append("\n")
                    list_en.append(text_en)
        # 各リストをそれぞれ一意のカラム名でDataFrame化
        df_ja = pd.DataFrame({'ja': list_ja})
        df_en = pd.DataFrame({'en': list_en})
        # インデックスで揃えて結合（列名がユニークなので再インデックスエラーを回避）
        df_m = pd.concat([df_ja, df_en], axis=1)
        # 全ページ分を連結する際はインデックスを振り直す
        df = pd.concat([df, df_m], ignore_index=True)
        
    df.to_excel(output_xlsx, index=False)

    print("処理が完了しました")


st.title("対訳ファイル作成")
st.write("和文と英文の統合報告書PDFから対訳ファイルのエクセルを作成します")

uploaded_file = st.file_uploader("和文（PDFファイル）をアップロード", type="pdf")
if uploaded_file is not None:
    file_name = os.path.splitext(uploaded_file.name)[0]  # アップロードされたファイル名を取得
    st.success("和文ファイルがアップロードされました。")
    doc_ja = pymupdf.open(stream=uploaded_file.read(), filetype="pdf")

uploaded_file = st.file_uploader("英文（PDFファイル）をアップロード", type="pdf")
if uploaded_file is not None:
    # file_name = os.path.splitext(uploaded_file.name)[0]  # アップロードされたファイル名を取得
    st.success("英文ファイルがアップロードされました。")
    doc_en = pymupdf.open(stream=uploaded_file.read(), filetype="pdf")
    
    try:
        output_xlsx = f"{file_name}_output.xlsx"  # アップロードされたファイル名をベースに出力ファイル名を作成
        extract_paragraphs_to_file(doc_ja, doc_en, output_xlsx)
        st.success(f"対訳ファイルが作成されました: {output_xlsx}")

        st.success("ダウンロードボタンを押してください")        
        with open(output_xlsx, "rb") as file:
            xlsx_data = file.read()
            # ダウンロードボタンを作成
            st.download_button(
                label="Excelをダウンロード",
                data=xlsx_data,
                file_name=file_name+"_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                on_click=None # 再実行を無視する設定（コールバックは不要）
            )
        print(f"ダウンロードしました。")

        
    except Exception as e:
        st.error(f"ファイルの読み込み中にエラーが発生しました: {e}")
else:
    st.info("ファイルをアップロードしてください。")
