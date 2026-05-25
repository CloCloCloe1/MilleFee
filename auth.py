from __future__ import annotations

import hashlib
import json
import os
import secrets
from datetime import datetime, timezone
from pathlib import Path


USER_STORE = Path(".app_users.json")
DEFAULT_ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "admin")
DEFAULT_ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "admin123")


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def hash_password(password: str, salt: str | None = None) -> str:
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 200_000)
    return f"pbkdf2_sha256${salt}${digest.hex()}"


def verify_password(password: str, password_hash: str) -> bool:
    try:
        algorithm, salt, expected = password_hash.split("$", 2)
    except ValueError:
        return False
    if algorithm != "pbkdf2_sha256":
        return False
    return secrets.compare_digest(hash_password(password, salt), f"{algorithm}${salt}${expected}")


def ensure_user_store() -> None:
    if USER_STORE.exists():
        return
    users = {
        DEFAULT_ADMIN_USERNAME: {
            "password_hash": hash_password(DEFAULT_ADMIN_PASSWORD),
            "role": "admin",
            "created_at": _now(),
            "updated_at": _now(),
            "must_change_password": DEFAULT_ADMIN_PASSWORD == "admin123",
        }
    }
    save_users(users)


def load_users() -> dict:
    ensure_user_store()
    return json.loads(USER_STORE.read_text(encoding="utf-8"))


def save_users(users: dict) -> None:
    USER_STORE.write_text(json.dumps(users, indent=2, sort_keys=True), encoding="utf-8")


def authenticate(username: str, password: str) -> dict | None:
    users = load_users()
    user = users.get(username.strip())
    if not user:
        return None
    if not verify_password(password, user.get("password_hash", "")):
        return None
    return {"username": username.strip(), "role": user.get("role", "user"), "must_change_password": user.get("must_change_password", False)}


def list_users() -> list[dict]:
    users = load_users()
    return [
        {
            "Username": username,
            "Role": data.get("role", "user"),
            "Created At": data.get("created_at", ""),
            "Updated At": data.get("updated_at", ""),
            "Must Change Password": bool(data.get("must_change_password", False)),
        }
        for username, data in sorted(users.items())
    ]


def create_user(username: str, password: str, role: str = "user") -> tuple[bool, str]:
    username = username.strip()
    if not username:
        return False, "Username is required."
    if len(password) < 6:
        return False, "Password must be at least 6 characters."
    if role not in {"admin", "user"}:
        return False, "Role must be admin or user."
    users = load_users()
    if username in users:
        return False, "Username already exists."
    users[username] = {
        "password_hash": hash_password(password),
        "role": role,
        "created_at": _now(),
        "updated_at": _now(),
        "must_change_password": False,
    }
    save_users(users)
    return True, "User created."


def delete_user(username: str, current_username: str) -> tuple[bool, str]:
    username = username.strip()
    if username == current_username:
        return False, "You cannot delete your own account while logged in."
    users = load_users()
    if username not in users:
        return False, "User does not exist."
    admin_count = sum(1 for data in users.values() if data.get("role") == "admin")
    if users[username].get("role") == "admin" and admin_count <= 1:
        return False, "At least one admin account is required."
    del users[username]
    save_users(users)
    return True, "User deleted."


def reset_password(username: str, new_password: str) -> tuple[bool, str]:
    username = username.strip()
    if len(new_password) < 6:
        return False, "Password must be at least 6 characters."
    users = load_users()
    if username not in users:
        return False, "User does not exist."
    users[username]["password_hash"] = hash_password(new_password)
    users[username]["updated_at"] = _now()
    users[username]["must_change_password"] = False
    save_users(users)
    return True, "Password updated."


def change_own_password(username: str, old_password: str, new_password: str) -> tuple[bool, str]:
    if len(new_password) < 6:
        return False, "New password must be at least 6 characters."
    user = authenticate(username, old_password)
    if not user:
        return False, "Current password is incorrect."
    return reset_password(username, new_password)
