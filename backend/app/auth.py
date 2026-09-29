import hashlib
import hmac
import re
import secrets
from datetime import datetime, timedelta, timezone

from fastapi import APIRouter, Depends, HTTPException, Request, Response
from pydantic import BaseModel, Field, field_validator
from sqlalchemy import select
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session

from .db import get_db
from .models import Project, SessionToken, User, now

router = APIRouter(prefix="/api")
COOKIE_NAME = "lukas_session"
SESSION_DAYS = 30
EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


class Credentials(BaseModel):
    email: str = Field(min_length=3, max_length=320)
    password: str = Field(min_length=8, max_length=128)

    @field_validator("email")
    @classmethod
    def normalize_email(cls, value: str) -> str:
        email = value.strip().casefold()
        if not EMAIL_PATTERN.fullmatch(email):
            raise ValueError("Укажите корректный адрес почты")
        return email


class AccountOut(BaseModel):
    id: str
    email: str | None


def password_digest(password: str) -> str:
    salt = secrets.token_bytes(16)
    digest = hashlib.scrypt(password.encode("utf-8"), salt=salt, n=2**14, r=8, p=1)
    return f"scrypt$16384${salt.hex()}${digest.hex()}"


def verify_password(password: str, stored: str | None) -> bool:
    if not stored:
        return False
    try:
        algorithm, cost, salt_hex, expected_hex = stored.split("$", 3)
        if algorithm != "scrypt" or cost != "16384":
            return False
        actual = hashlib.scrypt(password.encode("utf-8"), salt=bytes.fromhex(salt_hex), n=2**14, r=8, p=1)
        return hmac.compare_digest(actual, bytes.fromhex(expected_hex))
    except (ValueError, TypeError):
        return False


def session_id(token: str) -> str:
    return hashlib.sha256(token.encode("ascii")).hexdigest()


def current_user(request: Request, db: Session = Depends(get_db)) -> User:
    token = request.cookies.get(COOKIE_NAME)
    if not token:
        raise HTTPException(status_code=401, detail="Войдите в аккаунт")
    session = db.get(SessionToken, session_id(token))
    if session is None:
        raise HTTPException(status_code=401, detail="Сессия завершена. Войдите снова")
    expires_at = session.expires_at
    if expires_at.tzinfo is None:
        expires_at = expires_at.replace(tzinfo=timezone.utc)
    if expires_at <= datetime.now(timezone.utc):
        db.delete(session)
        db.commit()
        raise HTTPException(status_code=401, detail="Сессия завершена. Войдите снова")
    user = db.get(User, session.user_id)
    if user is None:
        raise HTTPException(status_code=401, detail="Сессия завершена. Войдите снова")
    return user


def start_session(response: Response, request: Request, db: Session, user: User) -> None:
    token = secrets.token_urlsafe(32)
    db.add(SessionToken(id=session_id(token), user_id=user.id, expires_at=now() + timedelta(days=SESSION_DAYS)))
    db.commit()
    response.set_cookie(
        COOKIE_NAME,
        token,
        max_age=SESSION_DAYS * 24 * 60 * 60,
        httponly=True,
        secure=request.url.scheme == "https" or request.headers.get("x-forwarded-proto") == "https",
        samesite="lax",
        path="/api",
    )
    response.headers["Cache-Control"] = "no-store"


@router.post("/auth/register", response_model=AccountOut, status_code=201)
def register(body: Credentials, request: Request, response: Response, db: Session = Depends(get_db)):
    if db.scalar(select(User.id).where(User.email == body.email)):
        raise HTTPException(status_code=409, detail="Аккаунт с этой почтой уже существует")
    user = User(email=body.email, password_hash=password_digest(body.password))
    db.add(user)
    try:
        db.flush()
        db.add(Project(name="Презентации", owner_user_id=user.id))
        db.commit()
    except IntegrityError as exc:
        db.rollback()
        raise HTTPException(status_code=409, detail="Аккаунт с этой почтой уже существует") from exc
    start_session(response, request, db, user)
    return AccountOut(id=user.id, email=user.email)


@router.post("/auth/login", response_model=AccountOut)
def login(body: Credentials, request: Request, response: Response, db: Session = Depends(get_db)):
    user = db.scalar(select(User).where(User.email == body.email))
    if user is None or not verify_password(body.password, user.password_hash):
        raise HTTPException(status_code=401, detail="Неверная почта или пароль")
    start_session(response, request, db, user)
    return AccountOut(id=user.id, email=user.email)


@router.get("/auth/me", response_model=AccountOut)
def me(user: User = Depends(current_user)):
    return AccountOut(id=user.id, email=user.email)


@router.post("/auth/logout", status_code=204)
def logout(request: Request, response: Response, db: Session = Depends(get_db)):
    token = request.cookies.get(COOKIE_NAME)
    if token:
        session = db.get(SessionToken, session_id(token))
        if session:
            db.delete(session)
            db.commit()
    response.delete_cookie(COOKIE_NAME, path="/api")
    response.headers["Cache-Control"] = "no-store"


@router.get("/workspace")
def workspace(user: User = Depends(current_user), db: Session = Depends(get_db)):
    project = db.scalar(select(Project).where(Project.owner_user_id == user.id).order_by(Project.created_at))
    if project is None:
        project = Project(name="Презентации", owner_user_id=user.id)
        db.add(project)
        db.commit()
        db.refresh(project)
    return {"id": project.id, "name": project.name, "created_at": project.created_at}
