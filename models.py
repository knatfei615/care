"""Database models and extension instances."""

from __future__ import annotations

from datetime import datetime, timezone

from flask_login import LoginManager, UserMixin
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import check_password_hash, generate_password_hash

db = SQLAlchemy()
login_manager = LoginManager()
login_manager.login_view = "auth.login"
login_manager.login_message = "请先登录。"


class User(db.Model, UserMixin):
    __tablename__ = "users"

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    display_name = db.Column(db.String(80), default="")
    role = db.Column(db.String(20), default="user")
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    def set_password(self, password: str) -> None:
        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
        return check_password_hash(self.password_hash, password)

    @property
    def is_admin(self) -> bool:
        return self.role == "admin"


@login_manager.user_loader
def _load_user(user_id: str) -> User | None:
    return db.session.get(User, int(user_id))


def init_db(app) -> None:
    """Create tables and seed admin user if configured."""
    with app.app_context():
        db.create_all()
        _seed_admin(app)


def _seed_admin(app) -> None:
    admin_username = app.config.get("ADMIN_USERNAME")
    admin_password = app.config.get("ADMIN_PASSWORD")
    if not admin_username or not admin_password:
        return
    if User.query.filter_by(username=admin_username).first():
        return
    admin = User(
        username=admin_username,
        display_name="管理员",
        role="admin",
    )
    admin.set_password(admin_password)
    db.session.add(admin)
    db.session.commit()
