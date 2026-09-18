import io
import os
import tempfile
from datetime import date
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4
from reportlab.lib.pagesizes import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
import pandas as pd
import qrcode
import streamlit as st

# --- 日本語フォント（Noto Sans JP）の登録 ---
try:
  font_path = "NotoSansJP-Regular.ttf"
  if os.path.exists(font_path):
    pdfmetrics.registerFont(TTFont("JapaneseFont", font_path))
  else:
    alt_path = "C:\\Windows\\Fonts\\meiryo.ttc"
    if os.path.exists(alt_path):
      pdfmetrics.registerFont(TTFont("JapaneseFont", alt_path, subfontIndex=0))
except Exception as e:
  print(f"フォント登録エラー: {e}")

# ページ設定
st.set_page_config(
    page_title="法人会イベント管理システム", layout="wide"
)

st.title("🏛️ 法人会イベント・会員管理システム")
st.markdown(
    "ローカルWindows環境稼働版（スマホQR・FAXハイブリッド対応）"
)

# タブの作成（タブ4を追加）
tab1, tab2, tab3, tab4 = st.tabs(
    [
        "タブ1: 参加者ピックアップ＆はがきPDF作成",
        "タブ2: 最終リスト編集 ＆ コンビニ収納CSV出力",
        "タブ3: 領収書PDF発行",
        "タブ4: スマホQR回答状況の確認",
    ]
)

# ==========================================
# タブ1: 参加者ピックアップ ＆ はがきPDF作成
# ==========================================
with tab1:
  st.header("1. イベント参加者選定 ＆ はがき・リスト出力")

  uploaded_file = st.file_uploader(
      "DBからエクスポートした会員リスト（Excel）をアップロード",
      type=["xlsx", "xls"],
      key="tab1_file",
  )

  if uploaded_file is not None:
    try:
      df_members = pd.read_excel(uploaded_file)
      st.success(
          f"会員データを正常に読み込みました（総件数: {len(df_members)}件）"
      )

      df_members.columns = df_members.columns.str.strip()

      reg_col = None
      for col in df_members.columns:
        if "整理番号" in col:
          reg_col = col
          break

      if reg_col is None:
        st.error(
            "エクセル内に「整理番号」を含む列が見つかりません。現在の列名:"
            f" {list(df_members.columns)}"
        )
      else:
        df_members = df_members.rename(
            columns={reg_col: "整理番号 ※重複不可"}
        )

        df_members["整理番号 ※重複不可"] = (
            df_members["整理番号 ※重複不可"]
            .astype(str)
            .str.replace(r"\.0$", "", regex=True)
            .str.zfill(6)
        )

        columns_list = list(df_members.columns)
        address_col_name = (
            columns_list[5] if len(columns_list) > 5 else None
        )
        postal_col_name = columns_list[6] if len(columns_list) > 6 else None

        st.subheader("🔍 会員絞り込み")
        col1, col2 = st.columns(2)

        with col1:
          blocks = ["すべて"] + list(
              df_members["ブロック"].dropna().unique()
              if "ブロック" in df_members.columns
              else []
          )
          selected_block = st.selectbox("ブロックで絞り込み", blocks, key="t1_b")

        with col2:
          if (
              selected_block != "すべて"
              and "ブロック" in df_members.columns
          ):
            filtered_for_支部 = df_members[
                df_members["ブロック"] == selected_block
            ]
          else:
            filtered_for_支部 = df_members
          branches = ["すべて"] + list(
              filtered_for_支部["支部"].dropna().unique()
              if "支部" in filtered_for_支部.columns
              else []
          )
          selected_branch = st.selectbox("支部で絞り込み", branches, key="t1_br")

        df_filtered = df_members.copy()
        if (
            selected_block != "すべて"
            and "ブロック" in df_filtered.columns
        ):
          df_filtered = df_filtered[df_filtered["ブロック"] == selected_block]
        if (
            selected_branch != "すべて"
            and "支部" in df_filtered.columns
        ):
          df_filtered = df_filtered[
              df_filtered["支部"] == selected_branch
          ]

        st.subheader("⚙️ イベント設定")
        col_e1, col_e2 = st.columns(2)
        with col_e1:
          event_name = st.text_input(
              "イベント名（出力7フィールド目）", value="第X回 交流会", key="t1_ev"
          )
        with col_e2:
          fee = st.number_input(
              "参加費（円）（出力6フィールド目）", value=5000, step=500, key="t1_fe"
          )

        free_memo = st.text_input(
            "はがき用フリーメモ（事務局情報の下に挿入されます。空欄でもOK）",
            value="",
            key="t1_memo",
        )

        st.subheader("✅ 参加対象者の選択")
        df_filtered["選択"] = False
        edited_df = st.data_editor(
            df_filtered,
            column_config={
                "選択": st.column_config.CheckboxColumn("選択", default=False)
            },
            disabled=[c for c in df_filtered.columns if c != "選択"],
            hide_index=True,
        )

        selected_rows = edited_df[edited_df["選択"] == True].copy()
        st.info(f"現在選択されている会員数: {len(selected_rows)}件")

        if st.button("📥 参加者応募リスト（Excel）を作成"):
          selected_rows["参加費"] = fee
          selected_rows["イベント名"] = event_name

          output = io.BytesIO()
          with pd.ExcelWriter(output, engine="openpyxl") as writer:
            selected_rows.to_excel(writer, index=False, sheet_name="参加者リスト")
          excel_data = output.getvalue()

          st.download_button(
              label="💾 参加者応募リストExcelをダウンロード",
              data=excel_data,
              file_name="participant_application_list.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
          )

        st.subheader("🖨️ はがきPDF生成")
        github_base_url = st.text_input(
            "GitHub Pagesの公開URL",
            value="https://ark-field.github.io/chibanshi-iventsanka.app/",
        )

        if st.button("📄 選択会員分のはがきPDFを生成"):
          if len(selected_rows) == 0:
            st.warning("対象者が選択されていません。")
          else:
            pdf_buffer = io.BytesIO()
            postcard_width = 100 * mm
            postcard_height = 148 * mm

            c = canvas.Canvas(
                pdf_buffer, pagesize=(postcard_width, postcard_height)
            )

            for idx, row in selected_rows.iterrows():
              reg_no = str(row["整理番号 ※重複不可"])
              corp_name = (
                  str(row["法人名"]) if "法人名" in row else "法人名不明"
              )

              addr = (
                  str(row[address_col_name])
                  if address_col_name and pd.notna(row[address_col_name])
                  else ""
              )
              postal_code = (
                  str(row[postal_col_name])
                  if postal_col_name and pd.notna(row[postal_col_name])
                  else ""
              )

              qr_url = f"{github_base_url}?id={reg_no}"

              qr = qrcode.QRCode(box_size=2, border=1)
              qr.add_data(qr_url)
              qr.make(fit=True)
              qr_img = qr.make_image(
                  fill_color="black", back_color="white"
              )

              temp_qr_path = f"temp_qr_{reg_no}.png"
              qr_img.save(temp_qr_path)

              try:
                c.setFont("JapaneseFont", 10)
              except:
                c.setFont("Helvetica", 10)

              if postal_code:
                c.drawString(
                    35 * mm, 115 * mm, f"〒 {postal_code.replace('〒', '')}"
                )

              display_name = f"{corp_name} 様"
              c.drawCentredString(42 * mm, 102 * mm, display_name)

              if addr:
                c.drawString(23 * mm, 95 * mm, addr)

              try:
                c.setFont("JapaneseFont", 9)
              except:
                c.setFont("Helvetica", 9)

              msg_line1 = f"「{event_name}」の参加にご希望の場合は"
              msg_line2 = (
                  "下記に〇してFAXまたはQRコードから参加を"
                  "クリックしてください。"
              )
              msg_line3 = "      参加する 参加しない"
              msg_line4 = "千葉西法人会事務局"
              msg_line5 = (
                  "〒262-0031 千葉市花見川区武石町2-612-16 TEL：043-272-8567"
              )

              c.drawString(10 * mm, 50 * mm, msg_line1)
              c.drawString(2.5 * mm, 45.2 * mm, msg_line2)

              try:
                c.setFont("JapaneseFont", 10)
              except:
                c.setFont("Helvetica", 10)
              c.drawString(10 * mm, 36 * mm, msg_line3)

              try:
                c.setFont("JapaneseFont", 8)
              except:
                c.setFont("Helvetica", 8)
              c.drawString(10 * mm, 26 * mm, msg_line4)
              c.drawString(10 * mm, 21 * mm, msg_line5)

              if free_memo:
                try:
                  c.setFont("JapaneseFont", 8)
                except:
                  c.setFont("Helvetica", 8)
                c.drawString(10 * mm, 16 * mm, f"メモ: {free_memo}")

              try:
                c.setFont("JapaneseFont", 11)
              except:
                c.setFont("Helvetica", 11)
              c.drawRightString(72 * mm, 28 * mm, f"ID: {reg_no}")

              c.drawImage(
                  temp_qr_path,
                  78 * mm,
                  24 * mm,
                  width=20 * mm,
                  height=20 * mm,
              )

              c.showPage()

              if os.path.exists(temp_qr_path):
                os.remove(temp_qr_path)

            c.save()
            pdf_buffer.seek(0)

            st.download_button(
                label="📥 はがきPDF（全員分まとめて）をダウンロード",
                data=pdf_buffer,
                file_name="postcard_tags.pdf",
                mime="application/pdf",
            )
            st.success("はがきPDFの生成が完了しました！")

    except Exception as e:
      st.error(f"ファイルの読み込み中にエラーが発生しました: {e}")

# ==========================================
# タブ2: 最終リスト編集 ＆ コンビニ収納CSV出力
# ==========================================
with tab2:
  st.header("2. 最終参加者リスト編集 ＆ コンビニ収納用WEB-EB CSV出力")
  st.markdown(
      "スマホ回答分およびFAX回答分をまとめた最終リストを読み込み・編集し、コンビニ収納用CSVを出力します。"
  )

  uploaded_file_tab2 = st.file_uploader(
      "最終参加者リスト（Excel）をアップロード",
      type=["xlsx", "xls"],
      key="tab2_file",
  )

  if uploaded_file_tab2 is not None:
    df_tab2 = pd.read_excel(uploaded_file_tab2)

    st.subheader("📝 リストの最終確認・手動調整（FAX分等の追加・修正）")
    edited_tab2 = st.data_editor(df_tab2, num_rows="dynamic", hide_index=True)

    if st.button("📤 コンビニ収納用WEB-EB CSVを生成", key="t2_csv_btn"):
      csv_data = edited_tab2.to_csv(
          index=False, encoding="cp932", errors="replace"
      )

      st.download_button(
          label="💾 WEB-EB用 CSVファイルをダウンロード",
          data=csv_data,
          file_name="convenience_store_web_eb.csv",
          mime="text/csv",
          key="t2_dl_btn",
      )
      st.success("CSVファイルを生成しました（Shift_JIS形式）")

# ==========================================
# タブ3: 領収書PDF発行（A4 2段切り取りタイプ・ブルー基調・社印欄なし）
# ==========================================
with tab3:
  st.header("3. 領収書PDF発行（A4用紙上下2段・切り取り形式）")
  st.markdown(
      "参加者リストのExcelをアップロードして、ブルー基調の綺麗な領収書をA4用紙1枚につき2件ずつ作成できます。"
  )

  uploaded_file_tab3 = st.file_uploader(
      "領収書発行用リスト（Excel）をアップロード",
      type=["xlsx", "xls"],
      key="tab3_file",
  )

  st.subheader("⚙️ 領収書共通設定")
  col_r1, col_r2 = st.columns(2)
  with col_r1:
    receipt_event_name = st.text_input(
        "但し書き（イベント名等）",
        value="第X回 交流会 参加費として",
        key="t3_ev",
    )
  with col_r2:
    receipt_date = st.date_input("発行日", value=date.today(), key="t3_dt")

  default_fee = st.number_input(
      "既定の金額（円）", value=5000, step=500, key="t3_fee"
  )

  if uploaded_file_tab3 is not None:
    df_rec = pd.read_excel(uploaded_file_tab3)
    df_rec.columns = df_rec.columns.str.strip()

    st.subheader("📝 発行対象の確認・金額調整")
    if "金額" not in df_rec.columns:
      df_rec["金額"] = default_fee

    edited_rec_df = st.data_editor(df_rec, num_rows="dynamic", hide_index=True)

    if st.button("📄 A4 2段切り取り式 領収書PDFを生成"):
      pdf_buffer = io.BytesIO()
      a4_w, a4_h = A4

      c = canvas.Canvas(pdf_buffer, pagesize=A4)

      records = list(edited_rec_df.iterrows())
      total_records = len(records)

      for i, (idx, row) in enumerate(records):
        pos_in_page = i % 2  # 0: 上段, 1: 下段

        if pos_in_page == 0:
          y_base = a4_h - 10 * mm
        else:
          y_base = (a4_h / 2.0) - 5 * mm

        corp_name = str(row["法人名"]) if "法人名" in row else "宛名不明"
        amt = (
            row["金額"]
            if "金額" in row and pd.notna(row["金額"])
            else default_fee
        )

        blue_primary = HexColor("#1A5276")
        blue_accent = HexColor("#AED6F1")
        char_dark = HexColor("#2C3E50")

        # 1. 上部の装飾ライン
        c.setStrokeColor(blue_primary)
        c.setLineWidth(1.5)
        c.line(20 * mm, y_base - 15 * mm, a4_w - 20 * mm, y_base - 15 * mm)

        # 2. 領収書タイトル
        c.setFillColor(blue_primary)
        try:
          c.setFont("JapaneseFont", 16)
        except:
          c.setFont("Helvetica-Bold", 16)
        c.drawString(20 * mm, y_base - 28 * mm, "領 収 書")

        # 3. 発行日
        c.setFillColor(char_dark)
        try:
          c.setFont("JapaneseFont", 9)
        except:
          c.setFont("Helvetica", 9)
        c.drawRightString(
            a4_w - 20 * mm,
            y_base - 25 * mm,
            f"発行日: {receipt_date.strftime('%Y年%m月%d日')}",
        )

        # 4. 宛名
        try:
          c.setFont("JapaneseFont", 13)
        except:
          c.setFont("Helvetica-Bold", 13)
        c.drawString(20 * mm, y_base - 42 * mm, f"{corp_name}  様")

        # 5. 金額ボックス
        c.setFillColor(HexColor("#F4F6F7"))
        c.setStrokeColor(blue_accent)
        c.rect(
            20 * mm,
            y_base - 65 * mm,
            a4_w - 40 * mm,
            16 * mm,
            fill=1,
            stroke=1,
        )

        c.setFillColor(blue_primary)
        try:
          c.setFont("JapaneseFont", 12)
        except:
          c.setFont("Helvetica", 12)
        c.drawString(25 * mm, y_base - 55 * mm, "金額:")
        c.drawRightString(
            a4_w - 25 * mm,
            y_base - 55 * mm,
            f"￥{int(amt):,} - (税込)",
        )

        # 6. 但し書き
        c.setFillColor(char_dark)
        try:
          c.setFont("JapaneseFont", 10)
        except:
          c.setFont("Helvetica", 10)
        c.drawString(
            20 * mm,
            y_base - 77 * mm,
            f"但し: {receipt_event_name}",
        )

        # 7. 発行元情報
        try:
          c.setFont("JapaneseFont", 10)
        except:
          c.setFont("Helvetica", 10)
        c.drawString(
            a4_w - 95 * mm, y_base - 28 * mm, "千葉西法人会事務局"
        )
        try:
          c.setFont("JapaneseFont", 8)
        except:
          c.setFont("Helvetica", 8)
        c.drawString(
            a4_w - 95 * mm,
            y_base - 34 * mm,
            "〒262-0031 千葉市花見川区武石町2-612-16",
        )
        c.drawString(
            a4_w - 95 * mm, y_base - 40 * mm, "TEL：043-272-8567"
        )

        # 8. 下部アンダーライン
        c.setStrokeColor(HexColor("#BDC3C7"))
        c.setLineWidth(0.5)
        c.line(20 * mm, y_base - 85 * mm, a4_w - 20 * mm, y_base - 85 * mm)

        # 9. 上段が終わったタイミングで、真ん中に切り取り線を描く
        if pos_in_page == 0 and i < total_records - 1:
          mid_y = a4_h / 2.0
          c.setStrokeColor(HexColor("#7F8C8D"))
          c.setLineWidth(1)
          c.setDash(4, 4)
          c.line(15 * mm, mid_y, a4_w - 15 * mm, mid_y)
          c.setDash()

          try:
            c.setFont("JapaneseFont", 8)
          except:
            c.setFont("Helvetica", 8)
          c.setFillColor(HexColor("#7F8C8D"))
          c.drawCentredString(
              a4_w / 2.0, mid_y + 2 * mm, "[ ここから切り取り ]"
          )

        if pos_in_page == 1 or i == total_records - 1:
          c.showPage()

      c.save()
      pdf_buffer.seek(0)

      st.download_button(
          label="📥 2段切り取り式 領収書PDFをダウンロード",
          data=pdf_buffer,
          file_name="receipts_2up.pdf",
          mime="application/pdf",
      )
      st.success("領収書PDFの生成が完了しました！")

# ==========================================
# タブ4: スマホQR回答状況の確認 ＆ CSV出力
# ==========================================
with tab4:
  st.header("4. スマホQR回答状況の確認 ＆ CSV出力")
  st.markdown(
      "スマホからQRコード経由で回答されたリアルタイムの集計結果をスプレッドシートから自動取得し、確認・CSV出力できます。"
  )

  sheet_csv_url = st.text_input(
      "Googleスプレッドシート「回答結果」のCSV公開リンク",
      value="",
      placeholder="https://docs.google.com/spreadsheets/d/.../export?format=csv&gid=0",
      key="t4_url",
  )
  st.info(
      "💡 スプレッドシートの「ファイル ＞ 共有 ＞ ウェブに公開」等でCSV形式（export?format=csv）にしたリンクを入力してください。"
  )

  df_tab4 = None
  if sheet_csv_url:
    if st.button("🔄 スプレッドシートから回答結果を読み込む", key="t4_load_btn"):
      try:
        df_tab4 = pd.read_csv(sheet_csv_url)
        st.session_state["df_qr_responses"] = df_tab4
        st.success(f"回答データを取得しました（件数: {len(df_tab4)}件）")
      except Exception as e:
        st.error(f"データの読み込みに失敗しました: {e}")

  if "df_qr_responses" in st.session_state:
    df_tab4 = st.session_state["df_qr_responses"]

  if df_tab4 is not None:
    st.subheader("📝 スマホ回答一覧の確認・編集")
    edited_tab4 = st.data_editor(df_tab4, num_rows="dynamic", hide_index=True)

    if st.button("📤 スマホ回答分のコンビニ収納CSVを生成", key="t4_csv_btn"):
      csv_data_t4 = edited_tab4.to_csv(
          index=False, encoding="cp932", errors="replace"
      )
      st.download_button(
          label="💾 スマホ回答用 WEB-EB CSVファイルをダウンロード",
          data=csv_data_t4,
          file_name="qr_responses_web_eb.csv",
          mime="text/csv",
          key="t4_dl_btn",
      )
      st.success("CSVファイルを生成しました（Shift_JIS形式）")
  else:
    st.markdown("上の入力欄にスプレッドシートのリンクを入力し、読み込みボタンを押してください。")