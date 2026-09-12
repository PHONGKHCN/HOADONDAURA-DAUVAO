import os
import json
import re
import sqlite3
from datetime import datetime
from io import BytesIO

from flask import Flask, render_template, request, redirect, url_for, flash, send_file, g, session
from flask_login import (
    LoginManager, UserMixin, login_user, logout_user,
    login_required, current_user
)
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from vn_number_to_words import so_thanh_chu

DB_PATH = os.path.join(os.path.dirname(__file__), "hoadon.db")

FONT_DIR = os.path.join(os.path.dirname(__file__), "static", "fonts")
pdfmetrics.registerFont(TTFont("VNSans", os.path.join(FONT_DIR, "DejaVuSans.ttf")))
pdfmetrics.registerFont(TTFont("VNSans-Bold", os.path.join(FONT_DIR, "DejaVuSans-Bold.ttf")))

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "doi-secret-key-nay-khi-deploy")

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"
login_manager.login_message = "Vui lòng đăng nhập để tiếp tục."


# ---------- DB helpers ----------

def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
        g.db.execute("PRAGMA journal_mode = WAL")
        g.db.execute("PRAGMA busy_timeout = 5000")
    return g.db


@app.teardown_appcontext
def close_db(exception=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    db = sqlite3.connect(DB_PATH)
    db.row_factory = sqlite3.Row
    db.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            full_name TEXT,
            role TEXT NOT NULL DEFAULT 'staff',  -- 'admin' | 'staff'
            created_at TEXT DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            loai TEXT NOT NULL,              -- 'dau_ra' | 'dau_vao'
            nhom TEXT,                       -- nhóm/khách hàng (VLXD, tạp hóa, v.v.)
            ky_hieu TEXT,
            so_hd TEXT NOT NULL,
            ngay_lap TEXT NOT NULL,          -- YYYY-MM-DD
            doi_tac TEXT NOT NULL,
            mst TEXT,
            mat_hang TEXT,
            doanh_so REAL NOT NULL DEFAULT 0,
            thue_suat REAL NOT NULL DEFAULT 10,
            tien_thue REAL NOT NULL DEFAULT 0,
            tong_cong REAL NOT NULL DEFAULT 0,
            ghi_chu TEXT,
            nguoi_tao_id INTEGER,
            created_at TEXT DEFAULT (datetime('now')),
            FOREIGN KEY (nguoi_tao_id) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS sales_invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nhom TEXT,                    -- nhóm/khách hàng (VLXD, tạp hóa, v.v.)
            so_hd TEXT,
            ngay_lap TEXT NOT NULL,
            seller_name TEXT NOT NULL,
            seller_address TEXT,
            seller_phone TEXT,
            buyer_name TEXT NOT NULL,
            buyer_address TEXT,
            tong_cong REAL NOT NULL DEFAULT 0,
            nguoi_tao_id INTEGER,
            created_at TEXT DEFAULT (datetime('now')),
            FOREIGN KEY (nguoi_tao_id) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS sales_invoice_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_id INTEGER NOT NULL,
            stt INTEGER NOT NULL,
            ten_hang TEXT NOT NULL,
            quy_cach TEXT,
            dvt TEXT,
            so_luong REAL NOT NULL DEFAULT 0,
            don_gia REAL NOT NULL DEFAULT 0,
            thanh_tien REAL NOT NULL DEFAULT 0,
            FOREIGN KEY (invoice_id) REFERENCES sales_invoices(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS partners (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            loai TEXT NOT NULL,           -- 'seller' | 'buyer'
            nhom TEXT,                    -- nhóm/khách hàng (vd: "Nguyen Van Do - VLXD")
            ten TEXT NOT NULL,
            dia_chi TEXT,
            sdt TEXT,
            mst TEXT,
            ghi_chu TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        );
        """
    )
    db.commit()

    # Migration an toàn: thêm cột 'nhom' nếu DB được tạo từ bản cũ chưa có cột này
    for table in ("invoices", "sales_invoices"):
        cols = [r["name"] for r in db.execute(f"PRAGMA table_info({table})").fetchall()]
        if "nhom" not in cols:
            db.execute(f"ALTER TABLE {table} ADD COLUMN nhom TEXT")
    db.commit()

    # Seed one admin account if no users exist yet
    cur = db.execute("SELECT COUNT(*) AS c FROM users")
    if cur.fetchone()["c"] == 0:
        db.execute(
            "INSERT INTO users (username, password_hash, full_name, role) VALUES (?, ?, ?, ?)",
            ("admin", generate_password_hash("admin123"), "Quản trị viên", "admin"),
        )
        db.commit()
        print("Đã tạo tài khoản mặc định: admin / admin123  (đổi mật khẩu ngay sau khi đăng nhập lần đầu)")
    db.close()


# ---------- User model ----------

class User(UserMixin):
    def __init__(self, row):
        self.id = row["id"]
        self.username = row["username"]
        self.full_name = row["full_name"]
        self.role = row["role"]

    @property
    def is_admin(self):
        return self.role == "admin"


@login_manager.user_loader
def load_user(user_id):
    db = get_db()
    row = db.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
    return User(row) if row else None


def admin_required(view):
    from functools import wraps

    @wraps(view)
    def wrapped(*args, **kwargs):
        if not current_user.is_authenticated or not current_user.is_admin:
            flash("Chỉ quản trị viên mới có quyền truy cập chức năng này.", "danger")
            return redirect(url_for("dashboard"))
        return view(*args, **kwargs)

    return wrapped


# ---------- Auth routes ----------

@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        db = get_db()
        row = db.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
        if row and check_password_hash(row["password_hash"], password):
            login_user(User(row))
            return redirect(url_for("dashboard"))
        flash("Sai tên đăng nhập hoặc mật khẩu.", "danger")
    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))


# ---------- Dashboard / reconciliation ----------

def month_bounds(period):
    """period format YYYY-MM -> (start_date, end_date) as strings YYYY-MM-DD"""
    year, month = map(int, period.split("-"))
    start = f"{year:04d}-{month:02d}-01"
    if month == 12:
        end = f"{year+1:04d}-01-01"
    else:
        end = f"{year:04d}-{month+1:02d}-01"
    return start, end


@app.route("/")
@login_required
def dashboard():
    period = request.args.get("period") or datetime.now().strftime("%Y-%m")
    start, end = month_bounds(period)
    db = get_db()

    def totals(loai):
        row = db.execute(
            """SELECT COALESCE(SUM(doanh_so),0) AS doanh_so,
                      COALESCE(SUM(tien_thue),0) AS tien_thue,
                      COALESCE(SUM(tong_cong),0) AS tong_cong,
                      COUNT(*) AS so_luong
               FROM invoices
               WHERE loai = ? AND ngay_lap >= ? AND ngay_lap < ?""",
            (loai, start, end),
        ).fetchone()
        return row

    dau_ra = totals("dau_ra")
    dau_vao = totals("dau_vao")
    chenh_lech = dau_ra["tien_thue"] - dau_vao["tien_thue"]

    return render_template(
        "dashboard.html",
        period=period,
        dau_ra=dau_ra,
        dau_vao=dau_vao,
        chenh_lech=chenh_lech,
    )


# ---------- Invoice CRUD ----------

@app.route("/invoices/bulk", methods=["GET", "POST"])
@login_required
def invoice_bulk():
    if request.method == "POST":
        raw = request.form.get("bulk_text", "")
        default_loai = request.form.get("default_loai", "")
        default_nhom = request.form.get("default_nhom", "").strip()
        lines = [l for l in raw.replace("\r\n", "\n").split("\n") if l.strip()]

        db = get_db()
        created = 0
        errors = []
        for idx, line in enumerate(lines, start=1):
            cols = [c.strip() for c in line.split("\t")]
            if len(cols) < 6:
                errors.append(f"Dòng {idx}: thiếu cột (cần ít nhất 6 cột)")
                continue
            try:
                loai_raw = cols[0].lower()
                if loai_raw in ("dau_vao", "đầu vào", "vao", "mua", "vào"):
                    loai = "dau_vao"
                elif loai_raw in ("dau_ra", "đầu ra", "ra", "ban", "bán"):
                    loai = "dau_ra"
                else:
                    loai = default_loai if default_loai in ("dau_vao", "dau_ra") else None
                if not loai:
                    errors.append(f"Dòng {idx}: không nhận diện được Loại ('{cols[0]}')")
                    continue

                ngay_raw = cols[1].strip()
                ngay_lap = None
                for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
                    try:
                        ngay_lap = datetime.strptime(ngay_raw, fmt).strftime("%Y-%m-%d")
                        break
                    except ValueError:
                        continue
                if not ngay_lap:
                    errors.append(f"Dòng {idx}: ngày không hợp lệ ('{ngay_raw}')")
                    continue

                doi_tac = cols[2]
                mst = cols[3] if len(cols) > 3 else ""
                mat_hang = cols[4] if len(cols) > 4 else ""
                doanh_so_raw = cols[5] if len(cols) > 5 else "0"
                doanh_so = float(re.sub(r"[^\d.\-]", "", doanh_so_raw.replace(",", "")) or 0)
                thue_suat = float(cols[6]) if len(cols) > 6 and cols[6] else 10
                so_hd = cols[7] if len(cols) > 7 else ""
                ky_hieu = cols[8] if len(cols) > 8 else ""
                ghi_chu = cols[9] if len(cols) > 9 else ""

                if not doi_tac:
                    errors.append(f"Dòng {idx}: thiếu tên đối tác")
                    continue

                tien_thue = round(doanh_so * thue_suat / 100, 0)
                tong_cong = doanh_so + tien_thue

                db.execute(
                    """INSERT INTO invoices
                       (loai, nhom, ky_hieu, so_hd, ngay_lap, doi_tac, mst, mat_hang,
                        doanh_so, thue_suat, tien_thue, tong_cong, ghi_chu, nguoi_tao_id)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (loai, default_nhom, ky_hieu, so_hd, ngay_lap, doi_tac, mst, mat_hang,
                     doanh_so, thue_suat, tien_thue, tong_cong, ghi_chu, current_user.id),
                )
                created += 1
            except Exception as e:
                errors.append(f"Dòng {idx}: lỗi xử lý ({e})")

        db.commit()
        if created:
            flash(f"Đã tạo {created} hóa đơn.", "success")
        if errors:
            flash("Một số dòng bị bỏ qua: " + " | ".join(errors[:10]) +
                  (f" (và {len(errors)-10} lỗi khác)" if len(errors) > 10 else ""), "danger")
        return redirect(url_for("invoice_list"))

    return render_template("invoice_bulk.html")


@app.route("/invoices")
@login_required
def invoice_list():
    loai = request.args.get("loai", "")
    period = request.args.get("period", "")
    nhom = request.args.get("nhom", "")
    q = request.args.get("q", "").strip()

    sql = """SELECT inv.*, u.full_name AS nguoi_tao_ten, u.username AS nguoi_tao_username
             FROM invoices inv LEFT JOIN users u ON inv.nguoi_tao_id = u.id WHERE 1=1"""
    params = []
    if loai in ("dau_ra", "dau_vao"):
        sql += " AND inv.loai = ?"
        params.append(loai)
    if nhom:
        sql += " AND inv.nhom = ?"
        params.append(nhom)
    if period:
        start, end = month_bounds(period)
        sql += " AND inv.ngay_lap >= ? AND inv.ngay_lap < ?"
        params += [start, end]
    if q:
        sql += " AND (inv.doi_tac LIKE ? OR inv.so_hd LIKE ? OR inv.mst LIKE ?)"
        like = f"%{q}%"
        params += [like, like, like]
    sql += " ORDER BY inv.ngay_lap DESC, inv.id DESC"

    db = get_db()
    invoices = db.execute(sql, params).fetchall()
    nhoms = db.execute(
        "SELECT DISTINCT nhom FROM invoices WHERE nhom IS NOT NULL AND nhom != '' ORDER BY nhom"
    ).fetchall()
    return render_template(
        "invoice_list.html", invoices=invoices, loai=loai, period=period, q=q,
        nhom=nhom, nhoms=nhoms,
    )


def parse_invoice_form(form):
    doanh_so = float(form.get("doanh_so") or 0)
    thue_suat = float(form.get("thue_suat") or 0)
    tien_thue = round(doanh_so * thue_suat / 100, 0)
    tong_cong = doanh_so + tien_thue
    return {
        "loai": form.get("loai"),
        "nhom": form.get("nhom", "").strip(),
        "ky_hieu": form.get("ky_hieu", "").strip(),
        "so_hd": form.get("so_hd", "").strip(),
        "ngay_lap": form.get("ngay_lap"),
        "doi_tac": form.get("doi_tac", "").strip(),
        "mst": form.get("mst", "").strip(),
        "mat_hang": form.get("mat_hang", "").strip(),
        "doanh_so": doanh_so,
        "thue_suat": thue_suat,
        "tien_thue": tien_thue,
        "tong_cong": tong_cong,
        "ghi_chu": form.get("ghi_chu", "").strip(),
    }


@app.route("/invoices/new", methods=["GET", "POST"])
@login_required
def invoice_new():
    if request.method == "POST":
        data = parse_invoice_form(request.form)
        db = get_db()
        db.execute(
            """INSERT INTO invoices
               (loai, nhom, ky_hieu, so_hd, ngay_lap, doi_tac, mst, mat_hang,
                doanh_so, thue_suat, tien_thue, tong_cong, ghi_chu, nguoi_tao_id)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                data["loai"], data["nhom"], data["ky_hieu"], data["so_hd"], data["ngay_lap"],
                data["doi_tac"], data["mst"], data["mat_hang"], data["doanh_so"],
                data["thue_suat"], data["tien_thue"], data["tong_cong"],
                data["ghi_chu"], current_user.id,
            ),
        )
        db.commit()
        flash("Đã lưu hóa đơn.", "success")
        return redirect(url_for("invoice_list", loai=data["loai"]))
    return render_template("invoice_form.html", invoice=None)


@app.route("/invoices/<int:invoice_id>/edit", methods=["GET", "POST"])
@login_required
def invoice_edit(invoice_id):
    db = get_db()
    invoice = db.execute("SELECT * FROM invoices WHERE id = ?", (invoice_id,)).fetchone()
    if not invoice:
        flash("Không tìm thấy hóa đơn.", "danger")
        return redirect(url_for("invoice_list"))

    if request.method == "POST":
        data = parse_invoice_form(request.form)
        db.execute(
            """UPDATE invoices SET loai=?, nhom=?, ky_hieu=?, so_hd=?, ngay_lap=?, doi_tac=?,
               mst=?, mat_hang=?, doanh_so=?, thue_suat=?, tien_thue=?, tong_cong=?, ghi_chu=?
               WHERE id=?""",
            (
                data["loai"], data["nhom"], data["ky_hieu"], data["so_hd"], data["ngay_lap"],
                data["doi_tac"], data["mst"], data["mat_hang"], data["doanh_so"],
                data["thue_suat"], data["tien_thue"], data["tong_cong"],
                data["ghi_chu"], invoice_id,
            ),
        )
        db.commit()
        flash("Đã cập nhật hóa đơn.", "success")
        return redirect(url_for("invoice_list", loai=data["loai"]))

    return render_template("invoice_form.html", invoice=invoice)


@app.route("/invoices/<int:invoice_id>/delete", methods=["POST"])
@login_required
def invoice_delete(invoice_id):
    db = get_db()
    db.execute("DELETE FROM invoices WHERE id = ?", (invoice_id,))
    db.commit()
    flash("Đã xóa hóa đơn.", "success")
    return redirect(request.referrer or url_for("invoice_list"))


# ---------- Excel export (theo mẫu bảng kê GTGT) ----------

FONT_NAME = "Times New Roman"


def style_ledger_sheet(ws, title, period, rows):
    header_fill = PatternFill(start_color="2E5266", end_color="2E5266", fill_type="solid")
    header_font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
    title_font = Font(name=FONT_NAME, bold=True, size=13)
    sub_font = Font(name=FONT_NAME, italic=True, size=10)
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")

    headers = ["STT", "Ký hiệu HĐ", "Số HĐ", "Ngày lập", "Tên đối tác", "Mã số thuế",
               "Mặt hàng/Dịch vụ", "Doanh số chưa thuế (đ)", "Thuế suất (%)",
               "Tiền thuế GTGT (đ)", "Tổng cộng (đ)", "Ghi chú"]
    widths = [5, 10, 8, 11, 22, 14, 22, 16, 10, 15, 16, 14]

    ws.merge_cells("A1:L1")
    ws["A1"] = title
    ws["A1"].font = title_font
    ws["A1"].alignment = center
    ws.merge_cells("A2:L2")
    ws["A2"] = f"Kỳ: {period}"
    ws["A2"].font = sub_font
    ws["A2"].alignment = center

    hr = 4
    for i, h in enumerate(headers):
        c = ws.cell(row=hr, column=1 + i, value=h)
        c.font = header_font
        c.fill = header_fill
        c.border = border
        c.alignment = center

    r = hr + 1
    for idx, inv in enumerate(rows, start=1):
        values = [
            idx, inv["ky_hieu"], inv["so_hd"], inv["ngay_lap"], inv["doi_tac"], inv["mst"],
            inv["mat_hang"], inv["doanh_so"], inv["thue_suat"], inv["tien_thue"],
            inv["tong_cong"], inv["ghi_chu"],
        ]
        for i, val in enumerate(values):
            c = ws.cell(row=r, column=1 + i, value=val)
            c.border = border
            c.font = Font(name=FONT_NAME, size=10)
            if i in (0, 1, 2, 3, 5, 8):
                c.alignment = center
            elif i in (7, 9, 10):
                c.alignment = right
                c.number_format = "#,##0"
            else:
                c.alignment = left
        r += 1

    if not rows:
        r += 1  # avoid SUM over header only

    total_row = r
    ws.cell(row=total_row, column=7, value="TỔNG CỘNG").font = Font(name=FONT_NAME, bold=True, size=10)
    ws.cell(row=total_row, column=7).alignment = Alignment(horizontal="right")
    for col_letter, col_idx in [("H", 8), ("J", 10), ("K", 11)]:
        cell = ws.cell(row=total_row, column=col_idx,
                        value=f"=SUM({col_letter}{hr+1}:{col_letter}{total_row-1})")
        cell.font = Font(name=FONT_NAME, bold=True, size=10)
        cell.number_format = "#,##0"
        cell.alignment = right
    for col_idx in range(1, 13):
        ws.cell(row=total_row, column=col_idx).border = border
        if col_idx not in (8, 10, 11):
            ws.cell(row=total_row, column=col_idx).fill = PatternFill(
                start_color="D6E4F0", end_color="D6E4F0", fill_type="solid"
            )

    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65 + i) if i < 26 else "A"].width = w
    ws.freeze_panes = "A5"
    return total_row


@app.route("/export")
@login_required
def export_excel():
    period = request.args.get("period") or datetime.now().strftime("%Y-%m")
    start, end = month_bounds(period)
    db = get_db()

    dau_ra = db.execute(
        "SELECT * FROM invoices WHERE loai='dau_ra' AND ngay_lap >= ? AND ngay_lap < ? ORDER BY ngay_lap",
        (start, end),
    ).fetchall()
    dau_vao = db.execute(
        "SELECT * FROM invoices WHERE loai='dau_vao' AND ngay_lap >= ? AND ngay_lap < ? ORDER BY ngay_lap",
        (start, end),
    ).fetchall()

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Đầu ra (Bán ra)"
    r1 = style_ledger_sheet(ws1, "BẢNG KÊ HÓA ĐƠN GTGT - ĐẦU RA (BÁN RA)", period, dau_ra)

    ws2 = wb.create_sheet("Đầu vào (Mua vào)")
    r2 = style_ledger_sheet(ws2, "BẢNG KÊ HÓA ĐƠN GTGT - ĐẦU VÀO (MUA VÀO)", period, dau_vao)

    ws3 = wb.create_sheet("Đối chiếu")
    title_font = Font(name=FONT_NAME, bold=True, size=13)
    header_fill = PatternFill(start_color="2E5266", end_color="2E5266", fill_type="solid")
    header_font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
    thin = Side(style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center")
    left = Alignment(horizontal="left")
    right = Alignment(horizontal="right")

    ws3.merge_cells("A1:D1")
    ws3["A1"] = f"ĐỐI CHIẾU THUẾ GTGT ĐẦU VÀO - ĐẦU RA (Kỳ {period})"
    ws3["A1"].font = title_font
    ws3["A1"].alignment = center

    labels = [
        ("Chỉ tiêu", "Đầu ra", "Đầu vào", "Chênh lệch"),
        ("Tổng doanh số chưa thuế", f"='Đầu ra (Bán ra)'!H{r1}", f"='Đầu vào (Mua vào)'!H{r2}", "=B6-C6"),
        ("Tổng tiền thuế GTGT", f"='Đầu ra (Bán ra)'!J{r1}", f"='Đầu vào (Mua vào)'!J{r2}", "=B7-C7"),
        ("Tổng cộng (có thuế)", f"='Đầu ra (Bán ra)'!K{r1}", f"='Đầu vào (Mua vào)'!K{r2}", "=B8-C8"),
    ]
    for i, row in enumerate(labels):
        for j, val in enumerate(row):
            c = ws3.cell(row=5 + i, column=1 + j, value=val)
            c.border = border
            if i == 0:
                c.font = header_font
                c.fill = header_fill
                c.alignment = center
            else:
                c.font = Font(name=FONT_NAME, size=10, bold=(j == 0))
                c.alignment = left if j == 0 else right
                if j > 0:
                    c.number_format = "#,##0"
    ws3.merge_cells("A10:D10")
    ws3["A10"] = "→ Chênh lệch dương: thuế GTGT phải nộp thêm | Chênh lệch âm: được khấu trừ/chuyển kỳ sau"
    ws3["A10"].font = Font(name=FONT_NAME, italic=True, size=9)
    for col, w in zip("ABCD", [28, 18, 18, 18]):
        ws3.column_dimensions[col].width = w

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"Bang_ke_hoa_don_{period}.xlsx"
    return send_file(
        buf, as_attachment=True, download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------- User management (admin only) ----------

@app.route("/users")
@login_required
@admin_required
def user_list():
    db = get_db()
    users = db.execute("SELECT * FROM users ORDER BY id").fetchall()
    return render_template("user_list.html", users=users)


@app.route("/users/new", methods=["GET", "POST"])
@login_required
@admin_required
def user_new():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        full_name = request.form.get("full_name", "").strip()
        role = request.form.get("role", "staff")
        db = get_db()
        exists = db.execute("SELECT id FROM users WHERE username=?", (username,)).fetchone()
        if exists:
            flash("Tên đăng nhập đã tồn tại.", "danger")
        else:
            db.execute(
                "INSERT INTO users (username, password_hash, full_name, role) VALUES (?,?,?,?)",
                (username, generate_password_hash(password), full_name, role),
            )
            db.commit()
            flash("Đã tạo tài khoản.", "success")
            return redirect(url_for("user_list"))
    return render_template("user_form.html")


@app.route("/users/<int:user_id>/delete", methods=["POST"])
@login_required
@admin_required
def user_delete(user_id):
    if user_id == current_user.id:
        flash("Không thể tự xóa tài khoản đang đăng nhập.", "danger")
        return redirect(url_for("user_list"))
    db = get_db()
    db.execute("DELETE FROM users WHERE id = ?", (user_id,))
    db.commit()
    flash("Đã xóa tài khoản.", "success")
    return redirect(url_for("user_list"))


# ---------- Danh mục đối tác (bên bán / bên mua đã lưu) ----------

@app.route("/partners")
@login_required
def partner_list():
    nhom = request.args.get("nhom", "")
    db = get_db()
    sql = "SELECT * FROM partners WHERE 1=1"
    params = []
    if nhom:
        sql += " AND nhom = ?"
        params.append(nhom)
    sql += " ORDER BY nhom, loai, ten"
    partners = db.execute(sql, params).fetchall()
    nhoms = db.execute(
        "SELECT DISTINCT nhom FROM partners WHERE nhom IS NOT NULL AND nhom != '' ORDER BY nhom"
    ).fetchall()
    return render_template("partner_list.html", partners=partners, nhoms=nhoms, nhom=nhom)


@app.route("/partners/new", methods=["GET", "POST"])
@login_required
def partner_new():
    if request.method == "POST":
        f = request.form
        db = get_db()
        db.execute(
            """INSERT INTO partners (loai, nhom, ten, dia_chi, sdt, mst, ghi_chu)
               VALUES (?,?,?,?,?,?,?)""",
            (f.get("loai"), f.get("nhom", "").strip(), f.get("ten", "").strip(),
             f.get("dia_chi", "").strip(), f.get("sdt", "").strip(),
             f.get("mst", "").strip(), f.get("ghi_chu", "").strip()),
        )
        db.commit()
        flash("Đã lưu đối tác.", "success")
        return redirect(url_for("partner_list"))
    return render_template("partner_form.html", partner=None)


@app.route("/partners/<int:partner_id>/edit", methods=["GET", "POST"])
@login_required
def partner_edit(partner_id):
    db = get_db()
    partner = db.execute("SELECT * FROM partners WHERE id=?", (partner_id,)).fetchone()
    if not partner:
        flash("Không tìm thấy đối tác.", "danger")
        return redirect(url_for("partner_list"))
    if request.method == "POST":
        f = request.form
        db.execute(
            """UPDATE partners SET loai=?, nhom=?, ten=?, dia_chi=?, sdt=?, mst=?, ghi_chu=?
               WHERE id=?""",
            (f.get("loai"), f.get("nhom", "").strip(), f.get("ten", "").strip(),
             f.get("dia_chi", "").strip(), f.get("sdt", "").strip(),
             f.get("mst", "").strip(), f.get("ghi_chu", "").strip(), partner_id),
        )
        db.commit()
        flash("Đã cập nhật đối tác.", "success")
        return redirect(url_for("partner_list"))
    return render_template("partner_form.html", partner=partner)


@app.route("/partners/<int:partner_id>/delete", methods=["POST"])
@login_required
def partner_delete(partner_id):
    db = get_db()
    db.execute("DELETE FROM partners WHERE id=?", (partner_id,))
    db.commit()
    flash("Đã xóa đối tác.", "success")
    return redirect(url_for("partner_list"))


# ---------- Tạo hóa đơn bán hàng (in được, nhiều mặt hàng) ----------

@app.route("/sales")
@login_required
def sales_list():
    nhom = request.args.get("nhom", "")
    db = get_db()
    sql = """SELECT si.*, u.full_name AS nguoi_tao_ten, u.username AS nguoi_tao_username
             FROM sales_invoices si LEFT JOIN users u ON si.nguoi_tao_id = u.id WHERE 1=1"""
    params = []
    if nhom:
        sql += " AND si.nhom = ?"
        params.append(nhom)
    sql += " ORDER BY si.ngay_lap DESC, si.id DESC"
    invoices = db.execute(sql, params).fetchall()
    nhoms = db.execute(
        "SELECT DISTINCT nhom FROM sales_invoices WHERE nhom IS NOT NULL AND nhom != '' ORDER BY nhom"
    ).fetchall()
    return render_template("sales_list.html", invoices=invoices, nhom=nhom, nhoms=nhoms)


@app.route("/sales/new", methods=["GET", "POST"])
@login_required
def sales_new():
    db = get_db()
    if request.method == "POST":
        f = request.form
        ten_hang_list = f.getlist("ten_hang[]")
        quy_cach_list = f.getlist("quy_cach[]")
        dvt_list = f.getlist("dvt[]")
        so_luong_list = f.getlist("so_luong[]")
        don_gia_list = f.getlist("don_gia[]")

        items = []
        tong_cong = 0
        stt = 0
        for i in range(len(ten_hang_list)):
            ten = ten_hang_list[i].strip()
            if not ten:
                continue
            stt += 1
            so_luong = float(so_luong_list[i] or 0)
            don_gia = float(don_gia_list[i] or 0)
            thanh_tien = so_luong * don_gia
            tong_cong += thanh_tien
            items.append({
                "stt": stt,
                "ten_hang": ten,
                "quy_cach": quy_cach_list[i].strip(),
                "dvt": dvt_list[i].strip(),
                "so_luong": so_luong,
                "don_gia": don_gia,
                "thanh_tien": thanh_tien,
            })

        partners_json = json.dumps([dict(p) for p in db.execute("SELECT * FROM partners ORDER BY nhom, loai, ten").fetchall()])

        if not items:
            flash("Cần ít nhất 1 mặt hàng.", "danger")
            return render_template(
                "sales_form.html", invoice=None, items=[],
                today=datetime.now().strftime("%Y-%m-%d"), partners_json=partners_json,
            )

        nhom = f.get("nhom", "").strip()
        seller_name = f.get("seller_name", "").strip()
        seller_address = f.get("seller_address", "").strip()
        seller_phone = f.get("seller_phone", "").strip()
        buyer_name = f.get("buyer_name", "").strip()
        buyer_address = f.get("buyer_address", "").strip()

        cur = db.execute(
            """INSERT INTO sales_invoices
               (nhom, so_hd, ngay_lap, seller_name, seller_address, seller_phone,
                buyer_name, buyer_address, tong_cong, nguoi_tao_id)
               VALUES (?,?,?,?,?,?,?,?,?,?)""",
            (
                nhom, f.get("so_hd", "").strip(), f.get("ngay_lap"),
                seller_name, seller_address, seller_phone, buyer_name,
                buyer_address, tong_cong, current_user.id,
            ),
        )
        invoice_id = cur.lastrowid
        for it in items:
            db.execute(
                """INSERT INTO sales_invoice_items
                   (invoice_id, stt, ten_hang, quy_cach, dvt, so_luong, don_gia, thanh_tien)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (invoice_id, it["stt"], it["ten_hang"], it["quy_cach"], it["dvt"],
                 it["so_luong"], it["don_gia"], it["thanh_tien"]),
            )

        # Lưu đối tác mới vào danh mục nếu được tick và chưa tồn tại
        if f.get("save_seller") and seller_name:
            exists = db.execute(
                "SELECT id FROM partners WHERE loai='seller' AND ten=? AND IFNULL(nhom,'')=?",
                (seller_name, nhom),
            ).fetchone()
            if not exists:
                db.execute(
                    "INSERT INTO partners (loai, nhom, ten, dia_chi, sdt) VALUES ('seller',?,?,?,?)",
                    (nhom, seller_name, seller_address, seller_phone),
                )
        if f.get("save_buyer") and buyer_name:
            exists = db.execute(
                "SELECT id FROM partners WHERE loai='buyer' AND ten=? AND IFNULL(nhom,'')=?",
                (buyer_name, nhom),
            ).fetchone()
            if not exists:
                db.execute(
                    "INSERT INTO partners (loai, nhom, ten, dia_chi) VALUES ('buyer',?,?,?)",
                    (nhom, buyer_name, buyer_address),
                )
        db.commit()

        flash("Đã tạo hóa đơn bán hàng.", "success")
        return redirect(url_for("sales_view", invoice_id=invoice_id))

    partners_json = json.dumps([dict(p) for p in db.execute("SELECT * FROM partners ORDER BY nhom, loai, ten").fetchall()])
    return render_template(
        "sales_form.html", invoice=None, items=[],
        today=datetime.now().strftime("%Y-%m-%d"), partners_json=partners_json,
    )


@app.route("/sales/<int:invoice_id>")
@login_required
def sales_view(invoice_id):
    db = get_db()
    invoice = db.execute("SELECT * FROM sales_invoices WHERE id=?", (invoice_id,)).fetchone()
    if not invoice:
        flash("Không tìm thấy hóa đơn.", "danger")
        return redirect(url_for("sales_list"))
    items = db.execute(
        "SELECT * FROM sales_invoice_items WHERE invoice_id=? ORDER BY stt", (invoice_id,)
    ).fetchall()
    return render_template(
        "sales_view.html", invoice=invoice, items=items,
        so_tien_chu=so_thanh_chu(invoice["tong_cong"]),
    )


@app.route("/sales/<int:invoice_id>/delete", methods=["POST"])
@login_required
def sales_delete(invoice_id):
    db = get_db()
    db.execute("DELETE FROM sales_invoice_items WHERE invoice_id=?", (invoice_id,))
    db.execute("DELETE FROM sales_invoices WHERE id=?", (invoice_id,))
    db.commit()
    flash("Đã xóa hóa đơn.", "success")
    return redirect(url_for("sales_list"))


def build_sales_invoice_pdf(invoice, items):
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        topMargin=15 * mm, bottomMargin=15 * mm,
        leftMargin=15 * mm, rightMargin=15 * mm,
    )

    style_normal = ParagraphStyle("normal", fontName="VNSans", fontSize=10, leading=13)
    style_bold = ParagraphStyle("bold", fontName="VNSans-Bold", fontSize=11, leading=14)
    style_title = ParagraphStyle(
        "title", fontName="VNSans-Bold", fontSize=16, leading=20, alignment=TA_CENTER,
        spaceAfter=4,
    )
    style_center = ParagraphStyle("center", fontName="VNSans", fontSize=10, alignment=TA_CENTER)
    style_right = ParagraphStyle("right", fontName="VNSans", fontSize=10, alignment=TA_RIGHT)
    style_cell = ParagraphStyle("cell", fontName="VNSans", fontSize=9, leading=11)
    style_cell_bold = ParagraphStyle("cell_bold", fontName="VNSans-Bold", fontSize=9, leading=11)

    story = []

    story.append(Paragraph(invoice["seller_name"] or "", style_bold))
    if invoice["seller_address"]:
        story.append(Paragraph(f"Địa chỉ: {invoice['seller_address']}", style_normal))
    if invoice["seller_phone"]:
        story.append(Paragraph(f"Điện thoại: {invoice['seller_phone']}", style_normal))
    story.append(Spacer(1, 8))

    story.append(Paragraph("HÓA ĐƠN BÁN HÀNG", style_title))
    if invoice["so_hd"]:
        story.append(Paragraph(f"Số: {invoice['so_hd']}", style_center))
    story.append(Spacer(1, 6))

    story.append(Paragraph(f"Khách hàng: {invoice['buyer_name']}", style_normal))
    if invoice["buyer_address"]:
        story.append(Paragraph(f"Địa chỉ: {invoice['buyer_address']}", style_normal))
    story.append(Spacer(1, 10))

    header = ["STT", "Tên hàng", "Quy cách\nsản phẩm", "ĐVT", "Số\nlượng", "Đơn giá", "Thành tiền"]
    table_data = [[Paragraph(h.replace("\n", "<br/>"), style_cell_bold) for h in header]]

    for it in items:
        table_data.append([
            Paragraph(str(it["stt"]), style_cell),
            Paragraph(it["ten_hang"], style_cell),
            Paragraph(it["quy_cach"] or "", style_cell),
            Paragraph(it["dvt"] or "", style_cell),
            Paragraph(f"{it['so_luong']:,.0f}".rstrip("0").rstrip(".") if it["so_luong"] % 1 else f"{it['so_luong']:,.0f}", style_cell),
            Paragraph(f"{it['don_gia']:,.0f}", style_cell),
            Paragraph(f"{it['thanh_tien']:,.0f}", style_cell),
        ])

    table_data.append([
        Paragraph("", style_cell), Paragraph("", style_cell), Paragraph("", style_cell),
        Paragraph("", style_cell), Paragraph("", style_cell),
        Paragraph("TỔNG CỘNG", style_cell_bold),
        Paragraph(f"{invoice['tong_cong']:,.0f}", style_cell_bold),
    ])

    col_widths = [12*mm, 45*mm, 28*mm, 15*mm, 18*mm, 25*mm, 27*mm]
    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.HexColor("#666666")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2E5266")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("ALIGN", (3, 0), (4, -1), "CENTER"),
        ("ALIGN", (5, 0), (6, -1), "RIGHT"),
        ("SPAN", (0, -1), (4, -1)),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#D6E4F0")),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 10))

    so_chu = so_thanh_chu(invoice["tong_cong"])
    story.append(Paragraph(f"Thành tiền (bằng chữ): <i>{so_chu}</i>", style_normal))
    story.append(Spacer(1, 4))
    ngay_str = invoice["ngay_lap"]
    try:
        d = datetime.strptime(ngay_str, "%Y-%m-%d")
        ngay_str = f"Ngày {d.day} tháng {d.month} năm {d.year}"
    except Exception:
        pass
    story.append(Paragraph(ngay_str, style_right))
    story.append(Spacer(1, 20))

    sig_table = Table(
        [[Paragraph("Khách hàng", style_center), Paragraph("Người bán", style_center)],
         [Paragraph("(Ký, ghi rõ họ tên)", style_center), Paragraph("(Ký, ghi rõ họ tên)", style_center)]],
        colWidths=[85 * mm, 85 * mm],
    )
    sig_table.setStyle(TableStyle([
        ("TOPPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(sig_table)

    doc.build(story)
    buf.seek(0)
    return buf


@app.route("/sales/<int:invoice_id>/pdf")
@login_required
def sales_pdf(invoice_id):
    db = get_db()
    invoice = db.execute("SELECT * FROM sales_invoices WHERE id=?", (invoice_id,)).fetchone()
    if not invoice:
        flash("Không tìm thấy hóa đơn.", "danger")
        return redirect(url_for("sales_list"))
    items = db.execute(
        "SELECT * FROM sales_invoice_items WHERE invoice_id=? ORDER BY stt", (invoice_id,)
    ).fetchall()
    buf = build_sales_invoice_pdf(invoice, items)
    filename = f"HoaDon_{invoice['so_hd'] or invoice_id}.pdf"
    return send_file(buf, as_attachment=True, download_name=filename, mimetype="application/pdf")


if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
else:
    # Ensure DB exists when run under gunicorn (Railway)
    init_db()
