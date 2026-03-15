"""Authentication blueprint – login / register / logout."""

from __future__ import annotations

from flask import Blueprint, flash, redirect, render_template, request, url_for
from flask_login import login_required, login_user, logout_user

from models import User, db

auth_bp = Blueprint("auth", __name__)


@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "GET":
        return render_template("login.html")

    username = (request.form.get("username") or "").strip()
    password = request.form.get("password") or ""

    if not username or not password:
        flash("请输入用户名和密码。", "error")
        return render_template("login.html"), 400

    user = User.query.filter_by(username=username).first()
    if not user or not user.check_password(password):
        flash("用户名或密码错误。", "error")
        return render_template("login.html"), 401

    login_user(user, remember=True)
    next_page = request.args.get("next") or url_for("index")
    return redirect(next_page)


@auth_bp.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "GET":
        return render_template("register.html")

    username = (request.form.get("username") or "").strip()
    display_name = (request.form.get("display_name") or "").strip()
    password = request.form.get("password") or ""
    confirm = request.form.get("confirm") or ""

    if not username or not password:
        flash("用户名和密码不能为空。", "error")
        return render_template("register.html"), 400

    if len(password) < 4:
        flash("密码至少 4 个字符。", "error")
        return render_template("register.html"), 400

    if password != confirm:
        flash("两次输入的密码不一致。", "error")
        return render_template("register.html"), 400

    if User.query.filter_by(username=username).first():
        flash("该用户名已被使用。", "error")
        return render_template("register.html"), 409

    is_first_user = User.query.count() == 0
    user = User(
        username=username,
        display_name=display_name or username,
        role="admin" if is_first_user else "user",
    )
    user.set_password(password)
    db.session.add(user)
    db.session.commit()

    login_user(user, remember=True)
    flash("注册成功！" + ("你是第一位用户，已自动设为管理员。" if is_first_user else ""), "success")
    return redirect(url_for("index"))


@auth_bp.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("auth.login"))
