from flask import Flask, render_template, request
import pandas as pd
import unicodedata
import difflib

app = Flask(__name__)

EXCEL_FILE = "20260205_管工事積算用データベース.xlsx"

def normalize_text(text):
    if pd.isna(text):
        return ""
    return unicodedata.normalize("NFKC", str(text)).strip()

# 🔥 起動時に1回だけ読み込む（高速化）
df = pd.read_excel(EXCEL_FILE, header=2)
df.columns = [normalize_text(col) for col in df.columns]

# 正規化列を事前作成
df["費目_norm"] = df["費目"].apply(normalize_text)
df["詳細規格_norm"] = df["詳細規格"].apply(normalize_text)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/search", methods=["POST"])
def search():
    keyword = request.form["keyword"]
    normalized_keyword = normalize_text(keyword)

    filtered = df[
        df["費目_norm"].str.contains(normalized_keyword, case=False, na=False) |
        df["詳細規格_norm"].str.contains(normalized_keyword, case=False, na=False)
    ].copy()

    suggestions = []
    if filtered.empty:
        all_words = df["費目_norm"].unique().tolist()
        suggestions = difflib.get_close_matches(normalized_keyword, all_words, n=5)

    # ===== 表示整形 =====

    # 年月表示
    if "公表年月" in filtered.columns:
        filtered["公表年月"] = pd.to_datetime(filtered["公表年月"], errors="coerce").dt.strftime("%Y/%m")

    # 数値カンマ
    for col in ["公表価格", "設計価格(or見積もり)", "当時のGaia価格"]:
        if col in filtered.columns:
            filtered[col] = pd.to_numeric(filtered[col], errors="coerce").map(lambda x: f"{int(x):,}" if pd.notna(x) else "")

    # %表示
    for col in ["設計採用率", "公表価格の採用率"]:
        if col in filtered.columns:
            filtered[col] = pd.to_numeric(filtered[col], errors="coerce").map(lambda x: f"{x*100:.1f}%" if pd.notna(x) else "")

    filtered = filtered.drop(columns=["費目_norm", "詳細規格_norm"], errors="ignore")

    results = filtered.to_dict(orient="records")
    columns = filtered.columns.tolist()

    return render_template("results.html",
                           results=results,
                           columns=columns,
                           keyword=keyword,
                           suggestions=suggestions)

if __name__ == "__main__":
    app.run(debug=True)