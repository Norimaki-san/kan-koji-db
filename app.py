import os
from flask import Flask, render_template, request, redirect, url_for, session
import pandas as pd
import unicodedata

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "devkey")

USERNAME = "nitto"
PASSWORD = "8326084"

EXCEL_FILE = "20260205_管工事積算用データベース.xlsx"

def normalize_text(text):
    if pd.isna(text):
        return ""
    return unicodedata.normalize("NFKC", str(text)).strip()

# ===== 起動時にデータ読み込み =====
df = pd.read_excel(EXCEL_FILE, header=2)
df.columns = [normalize_text(col) for col in df.columns]

df["費目_norm"] = df["費目"].apply(normalize_text)
df["詳細規格_norm"] = df["詳細規格"].apply(normalize_text)

# ===== 公表年月変換関数 =====
def convert_excel_serial(value):
    try:
        value = int(value)
        base_date = pd.Timestamp("1899-12-30")
        real_date = base_date + pd.Timedelta(days=value)
        return f"{real_date.year}年{real_date.month}月"
    except:
        return ""

@app.route("/login", methods=["GET", "POST"])
def login():
    error = None

    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")

        if username == USERNAME and password == PASSWORD:
            session["logged_in"] = True
            return redirect(url_for("index"))
        else:
            error = "IDまたはパスワードが違います"

    return render_template("login.html", error=error)

@app.route("/")
def index():
    if not session.get("logged_in"):
        return redirect(url_for("login"))
    return render_template(
        "index.html",
        results=None,
        show_columns=False
    )

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.after_request
def add_header(response):
    response.headers["Cache-Control"] = "no-store"
    return response

@app.route("/search", methods=["GET"])
def search():

    if not session.get("logged_in"):
        return redirect(url_for("login"))

    keyword = request.args.get("keyword", "")
    selected_columns = request.args.getlist("display_columns")

    normalized_keyword = normalize_text(keyword)

    if keyword:
        filtered = df[
            df["費目_norm"].str.contains(normalized_keyword, case=False, na=False) |
            df["詳細規格_norm"].str.contains(normalized_keyword, case=False, na=False)
        ].copy()
    else:
        filtered = df.copy()

    # ===== 列選択 =====
    if selected_columns:
        filtered = filtered[selected_columns]
    else:
        selected_columns = df.columns.tolist()
        filtered = filtered[selected_columns]

    # ===== 公表年月変換（Excelシリアル値対応）=====
    if "公表年月" in filtered.columns:
        filtered["公表年月"] = filtered["公表年月"].apply(convert_excel_serial)

    # ===== 数値整形 =====
    for col in ["公表価格", "設計価格(or見積もり)", "当時のGaia価格"]:
        if col in filtered.columns:
            filtered[col] = pd.to_numeric(
                filtered[col], errors="coerce"
            ).map(lambda x: f"{int(x):,}" if pd.notna(x) else "")

    for col in ["設計採用率", "公表価格の採用率"]:
        if col in filtered.columns:
            filtered[col] = pd.to_numeric(
                filtered[col], errors="coerce"
            ).map(lambda x: f"{x*100:.1f}%" if pd.notna(x) else "")

    filtered = filtered.drop(columns=["費目_norm", "詳細規格_norm"], errors="ignore")

    page = int(request.args.get("page", 1))
    per_page = 50

    start = (page - 1) * per_page
    end = start + per_page

    total_count = len(filtered)
    paginated = filtered.iloc[start:end]

    return render_template(
        "index.html",
        results=paginated.values.tolist(),
        columns=paginated.columns.tolist(),
        all_columns=df.columns.tolist(),
        selected_columns=selected_columns,
        count=total_count,
        highlight_keyword=keyword,
        page=page,
        show_columns=True
    )

if __name__ == "__main__":
    app.run(debug=True)