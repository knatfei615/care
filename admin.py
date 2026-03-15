"""Admin blueprint – user management."""

from __future__ import annotations

from functools import wraps

from flask import Blueprint, flash, redirect, render_template, request, url_for
from flask_login import current_user, login_required

from models import User, db

admin_bp = Blueprint("admin", __name__, url_prefix="/admin")


def admin_required(fn):
    @wraps(fn)
    @login_required
    def wrapper(*args, **kwargs):
        if not current_user.is_admin:
            flash("需要管理员权限。", "error")
            return redirect(url_for("index"))
        return fn(*args, **kwargs)
    return wrapper


@admin_bp.route("/users")
@admin_required
def user_list():
    users = User.query.order_by(User.created_at.desc()).all()
    return render_template("admin.html", users=users)


@admin_bp.route("/users", methods=["POST"])
@admin_required
def create_user():
    username = (request.form.get("username") or "").strip()
    display_name = (request.form.get("display_name") or "").strip()
    password = request.form.get("password") or ""
    role = request.form.get("role", "user")

    if not username or not password:
        flash("用户名和密码不能为空。", "error")
        return redirect(url_for("admin.user_list"))

    if User.query.filter_by(username=username).first():
        flash("该用户名已被使用。", "error")
        return redirect(url_for("admin.user_list"))

    if role not in ("admin", "user"):
        role = "user"

    user = User(username=username, display_name=display_name or username, role=role)
    user.set_password(password)
    db.session.add(user)
    db.session.commit()
    flash(f"用户 {username} 创建成功。", "success")
    return redirect(url_for("admin.user_list"))


@admin_bp.route("/users/<int:user_id>/delete", methods=["POST"])
@admin_required
def delete_user(user_id: int):
    user = db.session.get(User, user_id)
    if not user:
        flash("用户不存在。", "error")
        return redirect(url_for("admin.user_list"))
    if user.id == current_user.id:
        flash("不能删除自己的账号。", "error")
        return redirect(url_for("admin.user_list"))

    db.session.delete(user)
    db.session.commit()
    flash(f"用户 {user.username} 已删除。", "success")
    return redirect(url_for("admin.user_list"))


@admin_bp.route("/users/<int:user_id>/reset-pw", methods=["POST"])
@admin_required
def reset_password(user_id: int):
    user = db.session.get(User, user_id)
    if not user:
        flash("用户不存在。", "error")
        return redirect(url_for("admin.user_list"))

    new_pw = request.form.get("password") or ""
    if len(new_pw) < 4:
        flash("密码至少 4 个字符。", "error")
        return redirect(url_for("admin.user_list"))

    user.set_password(new_pw)
    db.session.commit()
    flash(f"用户 {user.username} 的密码已重置。", "success")
    return redirect(url_for("admin.user_list"))
