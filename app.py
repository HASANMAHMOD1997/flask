import os
import io
import sqlite3
from flask import Flask, render_template, request, redirect, url_for, flash, send_file


app = Flask(__name__)
app.secret_key = "change-this-secret-key"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Allow overriding DB path via environment (for Render disk mount)
DB_PATH = os.environ.get("DATABASE_PATH") or os.path.join(BASE_DIR, "database.db")


def initialize_database() -> None:
    # Ensure parent directory exists when DB is on a mounted disk
    parent_dir = os.path.dirname(DB_PATH)
    if parent_dir and not os.path.exists(parent_dir):
        os.makedirs(parent_dir, exist_ok=True)

    connection = sqlite3.connect(DB_PATH)
    try:
        cursor = connection.cursor()
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                age INTEGER NOT NULL CHECK(age >= 0)
            );
            """
        )
        connection.commit()
    finally:
        connection.close()


def insert_user(name: str, age: int) -> None:
    connection = sqlite3.connect(DB_PATH)
    try:
        cursor = connection.cursor()
        cursor.execute(
            "INSERT INTO users (name, age) VALUES (?, ?)",
            (name, age),
        )
        connection.commit()
    finally:
        connection.close()


def get_user_by_id(user_id: int):
    connection = sqlite3.connect(DB_PATH)
    try:
        cursor = connection.cursor()
        cursor.execute("SELECT id, name, age FROM users WHERE id = ?", (user_id,))
        row = cursor.fetchone()
        return row
    finally:
        connection.close()


def update_user(user_id: int, name: str, age: int) -> None:
    connection = sqlite3.connect(DB_PATH)
    try:
        cursor = connection.cursor()
        cursor.execute(
            "UPDATE users SET name = ?, age = ? WHERE id = ?",
            (name, age, user_id),
        )
        connection.commit()
    finally:
        connection.close()


def delete_user(user_id: int) -> None:
    connection = sqlite3.connect(DB_PATH)
    try:
        cursor = connection.cursor()
        cursor.execute("DELETE FROM users WHERE id = ?", (user_id,))
        connection.commit()
    finally:
        connection.close()


@app.route("/", methods=["GET"]) 
def index():
    return render_template("index.html")


@app.route("/submit", methods=["POST"]) 
def submit():
    name = (request.form.get("name") or "").strip()
    age_raw = (request.form.get("age") or "").strip()

    if not name:
        flash("الاسم مطلوب")
        return redirect(url_for("index"))

    try:
        age = int(age_raw)
        if age < 0:
            raise ValueError("Age must be non-negative")
    except ValueError:
        flash("العمر يجب أن يكون رقمًا صحيحًا غير سالب")
        return redirect(url_for("index"))

    insert_user(name, age)
    return render_template("success.html", name=name, age=age)


@app.route("/list", methods=["GET"]) 
def list_users():
    q = (request.args.get("q") or "").strip()
    connection = sqlite3.connect(DB_PATH)
    try:
        cursor = connection.cursor()
        if q:
            cursor.execute(
                "SELECT id, name, age FROM users WHERE name LIKE ? ORDER BY id DESC",
                (f"%{q}%",),
            )
        else:
            cursor.execute("SELECT id, name, age FROM users ORDER BY id DESC")
        rows = cursor.fetchall()
    finally:
        connection.close()
    return render_template("list.html", users=rows, q=q)


@app.route("/export", methods=["GET"]) 
def export_users():
    q = (request.args.get("q") or "").strip()
    connection = sqlite3.connect(DB_PATH)
    try:
        cursor = connection.cursor()
        if q:
            cursor.execute(
                "SELECT id, name, age FROM users WHERE name LIKE ? ORDER BY id DESC",
                (f"%{q}%",),
            )
        else:
            cursor.execute("SELECT id, name, age FROM users ORDER BY id DESC")
        rows = cursor.fetchall()
    finally:
        connection.close()

    # Build Excel in-memory using openpyxl
    try:
        from openpyxl import Workbook
    except Exception:
        flash("يرجى تثبيت مكتبة openpyxl عبر requirements.txt")
        return redirect(url_for("list_users", q=q))

    wb = Workbook()
    ws = wb.active
    ws.title = "users"
    ws.append(["ID", "Name", "Age"])  # header
    for row in rows:
        ws.append([row[0], row[1], row[2]])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = "users.xlsx" if not q else f"users_search_{q}.xlsx"
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


@app.route("/edit/<int:user_id>", methods=["GET", "POST"]) 
def edit_user(user_id: int):
    existing = get_user_by_id(user_id)
    if not existing:
        flash("السجل غير موجود")
        return redirect(url_for("list_users"))

    if request.method == "POST":
        name = (request.form.get("name") or "").strip()
        age_raw = (request.form.get("age") or "").strip()

        if not name:
            flash("الاسم مطلوب")
            return redirect(url_for("edit_user", user_id=user_id))

        try:
            age = int(age_raw)
            if age < 0:
                raise ValueError("Age must be non-negative")
        except ValueError:
            flash("العمر يجب أن يكون رقمًا صحيحًا غير سالب")
            return redirect(url_for("edit_user", user_id=user_id))

        update_user(user_id, name, age)
        flash("تم تحديث السجل بنجاح")
        return redirect(url_for("list_users"))

    # GET
    return render_template("edit.html", user_id=existing[0], name=existing[1], age=existing[2])


@app.route("/delete/<int:user_id>", methods=["POST"]) 
def remove_user(user_id: int):
    existing = get_user_by_id(user_id)
    if not existing:
        flash("السجل غير موجود")
        return redirect(url_for("list_users"))
    delete_user(user_id)
    flash("تم حذف السجل")
    return redirect(url_for("list_users"))


if __name__ == "__main__":
    initialize_database()
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)


