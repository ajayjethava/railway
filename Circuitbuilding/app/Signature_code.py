#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════════════════╗
║         Digital Signature Verification System  •  v2.0.0                   ║
║         Single-file, industry-grade Python application                     ║
╚══════════════════════════════════════════════════════════════════════════════╝

Register and verify PDF digital signatures against a local certificate
registry backed by SQLite.  All operations are recorded to a dual-channel
logging system (application log + immutable audit trail).

Features
--------
  • Register users via their digitally-signed PDF certificate
  • Verify a PDF signer against the local registry
  • List / search / revoke registered users
  • Export the registry to a timestamped CSV file
  • Rotating file-based app + audit logs (NO_COLOR-aware ANSI output)
  • WAL-mode SQLite with per-operation rollback safety

Usage
-----
    python final.py

Dependency
----------
    pip install pyhanko
"""

from __future__ import annotations

import csv
import logging
import logging.handlers
import os
import shutil
import sqlite3
import sys
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Generator, List, Optional, Tuple

# ══════════════════════════════════════════════════════════════════════════════
# §1  METADATA
# ══════════════════════════════════════════════════════════════════════════════

__version__  = "2.0.0"
__app_name__ = "Digital Signature Verification System"
__author__   = "Your Organisation"


# ══════════════════════════════════════════════════════════════════════════════
# §2  CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════════

@dataclass(frozen=True)
class _Config:
    """Central, immutable configuration for the entire application.

    All paths are relative to the current working directory unless overridden
    via environment variables before import.
    """
    db_path       : Path = Path(os.getenv("DSS_DB",         "signature_users.db"))
    log_dir       : Path = Path(os.getenv("DSS_LOG_DIR",    "logs"))
    export_dir    : Path = Path(os.getenv("DSS_EXPORT_DIR", "exports"))
    max_log_bytes : int  = 5 * 1024 * 1024   # 5 MB per rotating file
    backup_count  : int  = 5                  # keep 5 rotated copies


CFG = _Config()


# ══════════════════════════════════════════════════════════════════════════════
# §3  ANSI COLOUR HELPERS
#     Respects the NO_COLOR environment variable and non-TTY output.
# ══════════════════════════════════════════════════════════════════════════════

_USE_COLOR: bool = sys.stdout.isatty() and os.getenv("NO_COLOR") is None


class _C:
    """Namespace of zero-dependency ANSI colour helpers."""

    RESET  = "\033[0m"
    BOLD   = "\033[1m"
    DIM    = "\033[2m"
    RED    = "\033[91m"
    GREEN  = "\033[92m"
    YELLOW = "\033[93m"
    CYAN   = "\033[96m"

    @classmethod
    def _w(cls, text: str, code: str) -> str:
        return f"{code}{text}{cls.RESET}" if _USE_COLOR else text

    @classmethod
    def red(cls, t: str)    -> str: return cls._w(t, cls.RED)
    @classmethod
    def green(cls, t: str)  -> str: return cls._w(t, cls.GREEN)
    @classmethod
    def yellow(cls, t: str) -> str: return cls._w(t, cls.YELLOW)
    @classmethod
    def cyan(cls, t: str)   -> str: return cls._w(t, cls.CYAN)
    @classmethod
    def bold(cls, t: str)   -> str: return cls._w(t, cls.BOLD)
    @classmethod
    def dim(cls, t: str)    -> str: return cls._w(t, cls.DIM)


# ══════════════════════════════════════════════════════════════════════════════
# §4  CUSTOM EXCEPTIONS
# ══════════════════════════════════════════════════════════════════════════════

class DSSError(Exception):
    """Root exception for the Digital Signature System."""

class CertificateError(DSSError):
    """Certificate extraction or validation failed."""

class DatabaseError(DSSError):
    """A database operation failed."""

class UserAlreadyExistsError(DSSError):
    """A certificate with this serial number is already registered."""

class FileValidationError(DSSError):
    """The supplied file path is invalid or unreadable."""


# ══════════════════════════════════════════════════════════════════════════════
# §5  LOGGING
#     Two rotating loggers:
#       dss.app   – general application events  (DEBUG and above)
#       dss.audit – security-relevant actions   (INFO and above)
# ══════════════════════════════════════════════════════════════════════════════

def _setup_logging() -> Tuple[logging.Logger, logging.Logger]:
    """Configure and return the application and audit loggers."""
    CFG.log_dir.mkdir(parents=True, exist_ok=True)

    def _handler(filename: Path, fmt: str) -> logging.handlers.RotatingFileHandler:
        h = logging.handlers.RotatingFileHandler(
            filename,
            maxBytes=CFG.max_log_bytes,
            backupCount=CFG.backup_count,
            encoding="utf-8",
        )
        h.setFormatter(logging.Formatter(fmt, datefmt="%Y-%m-%d %H:%M:%S"))
        return h

    app_log = logging.getLogger("dss.app")
    app_log.setLevel(logging.DEBUG)
    app_log.addHandler(_handler(
        CFG.log_dir / "dss.log",
        "[%(asctime)s] [%(levelname)-8s] %(message)s",
    ))

    aud_log = logging.getLogger("dss.audit")
    aud_log.setLevel(logging.INFO)
    aud_log.addHandler(_handler(
        CFG.log_dir / "audit.log",
        "[%(asctime)s] AUDIT | %(message)s",
    ))

    return app_log, aud_log


_log, _audit = _setup_logging()


# ══════════════════════════════════════════════════════════════════════════════
# §6  DATA MODELS
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class CertificateInfo:
    """Raw certificate metadata extracted from a signed PDF."""
    field_name    : str
    subject       : str
    issuer        : str
    serial_number : str


@dataclass
class RegisteredUser:
    """A user record stored in the local certificate registry."""
    id            : int
    username      : str
    serial_number : str
    subject       : str
    issuer        : str
    registered_at : str


@dataclass
class VerificationResult:
    """Outcome of a signature verification check.

    Attributes
    ----------
    status:
        One of ``"APPROVED"``, ``"INVALID"``, or ``"ERROR"``.
    """
    status        : str
    username      : Optional[str] = None
    serial_number : Optional[str] = None
    message       : Optional[str] = None


# ══════════════════════════════════════════════════════════════════════════════
# §7  DATABASE LAYER
# ══════════════════════════════════════════════════════════════════════════════

@contextmanager
def _db() -> Generator[sqlite3.Connection, None, None]:
    """Yield a WAL-mode SQLite connection with auto-commit / auto-rollback.

    ``sqlite3.Row`` is active as the row factory so columns are accessible
    by name.  Each call opens a fresh connection; nesting is safe.
    """
    conn = sqlite3.connect(str(CFG.db_path), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    try:
        yield conn
        conn.commit()
    except sqlite3.Error as exc:
        conn.rollback()
        _log.error("DB rollback — %s", exc)
        raise DatabaseError(str(exc)) from exc
    finally:
        conn.close()


def init_database() -> None:
    """Create all schema objects if they do not already exist."""
    with _db() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS users (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                username       TEXT    NOT NULL,
                serial_number  TEXT    UNIQUE NOT NULL,
                subject        TEXT,
                issuer         TEXT,
                registered_at  TEXT    NOT NULL
            );

            CREATE TABLE IF NOT EXISTS audit_log (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                action       TEXT    NOT NULL,
                detail       TEXT,
                performed_at TEXT    NOT NULL
            );
        """)
    _log.info("Database ready — %s", CFG.db_path)


def _write_audit(action: str, detail: str) -> None:
    """Persist an audit record in the DB table and the rotating audit log."""
    try:
        with _db() as conn:
            conn.execute(
                "INSERT INTO audit_log(action, detail, performed_at) VALUES (?,?,?)",
                (action, detail, datetime.now().isoformat()),
            )
        _audit.info("%-22s | %s", action, detail)
    except DatabaseError:
        _audit.error(
            "Failed to persist audit record — action=%s detail=%s", action, detail
        )


def _row_to_user(r: sqlite3.Row) -> RegisteredUser:
    """Map a database row to a ``RegisteredUser`` dataclass."""
    return RegisteredUser(
        id            = r["id"],
        username      = r["username"],
        serial_number = r["serial_number"],
        subject       = r["subject"] or "",
        issuer        = r["issuer"]  or "",
        registered_at = r["registered_at"],
    )


# ══════════════════════════════════════════════════════════════════════════════
# §8  CERTIFICATE EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════

def extract_signature_info(pdf_path: Path) -> CertificateInfo:
    """Extract the first embedded digital-signature certificate from a PDF.

    Parameters
    ----------
    pdf_path:
        Absolute or relative path to the signed PDF.

    Returns
    -------
    CertificateInfo
        Parsed certificate metadata.

    Raises
    ------
    FileValidationError
        ``pdf_path`` does not exist or is not a regular file.
    CertificateError
        No signature found, or the signer certificate is absent.
    DSSError
        ``pyhanko`` library is not installed.
    """
    if not pdf_path.is_file():
        raise FileValidationError(f"File not found: {pdf_path}")

    try:
        from pyhanko.pdf_utils.reader import PdfFileReader  # deferred import
    except ImportError as exc:
        raise DSSError(
            "pyhanko is not installed — run:  pip install pyhanko"
        ) from exc

    try:
        serials = []
        with pdf_path.open("rb") as fh:
            reader = PdfFileReader(fh)
            sigs   = list(reader.embedded_signatures)

        if not sigs:
            raise CertificateError("No digital signature found in the PDF.")
        def get_sig_time(sig):
            # PyPDF2 signature date fields can vary
            return (
                getattr(sig, "signing_time", None)
                or getattr(sig, "sig_date", None)
                or getattr(sig, "timestamp", None)
            )

        sigs = sorted(
            sigs,
            key=lambda s: get_sig_time(s) or 0,
            reverse=True
        )  
        sig  = sigs[0]
        cert = sig.signer_cert
        if cert is None:
            raise CertificateError(
                "Signer certificate is not embedded in the signature field."
            )

        info = CertificateInfo(
            field_name    = sig.field_name,
            subject       = cert.subject.human_friendly,
            issuer        = cert.issuer.human_friendly,
            serial_number = str(cert.serial_number),
        )
        _log.debug(
            "Certificate extracted from '%s' — serial=%s",
            pdf_path.name, info.serial_number,
        )
        return info

    except CertificateError:
        raise
    except Exception as exc:
        _log.exception("Unexpected error reading '%s'", pdf_path)
        raise CertificateError(f"Failed to read PDF signature: {exc}") from exc


def extract_all_signature_info(pdf_path: Path) -> list[CertificateInfo]:
    """
    Extract ALL embedded digital-signature certificates from a PDF.

    Returns
    -------
    list[CertificateInfo]
        List of all signer certificates found in the PDF.
    """

    if not pdf_path.is_file():
        raise FileValidationError(f"File not found: {pdf_path}")

    try:
        from pyhanko.pdf_utils.reader import PdfFileReader
    except ImportError as exc:
        raise DSSError(
            "pyhanko is not installed — run: pip install pyhanko"
        ) from exc

    try:
        with pdf_path.open("rb") as fh:
            reader = PdfFileReader(fh)
            sigs = list(reader.embedded_signatures)

        if not sigs:
            raise CertificateError("No digital signature found in the PDF.")

        cert_infos = []

        for sig in sigs:
            cert = sig.signer_cert
            if cert is None:
                continue

            info = CertificateInfo(
                field_name=sig.field_name,
                subject=cert.subject.human_friendly,
                issuer=cert.issuer.human_friendly,
                serial_number=str(cert.serial_number),
            )
            cert_infos.append(info)

        if not cert_infos:
            raise CertificateError("No valid signer certificates found.")

        return cert_infos

    except CertificateError:
        raise
    except Exception as exc:
        _log.exception("Error reading PDF signatures: %s", pdf_path)
        raise CertificateError(f"Failed to read PDF signature: {exc}") from exc
# ══════════════════════════════════════════════════════════════════════════════
# §9  SERVICE LAYER
# ══════════════════════════════════════════════════════════════════════════════

def register_user(username: str, pdf_path: Path) -> RegisteredUser:
    """Register a new user by storing their certificate from a signed PDF.

    Parameters
    ----------
    username:
        Human-readable display name — must be non-empty.
    pdf_path:
        Path to a PDF signed with the user's private key.

    Returns
    -------
    RegisteredUser
        The newly created registry entry.

    Raises
    ------
    ValueError
        ``username`` is blank.
    UserAlreadyExistsError
        Another user with the same certificate serial already exists.
    FileValidationError, CertificateError, DatabaseError
        Propagated from lower layers.
    """
    username = username.strip()
    if not username:
        raise ValueError("Username cannot be empty.")

    cert = extract_signature_info(pdf_path)

    with _db() as conn:
        existing = conn.execute(
            "SELECT username FROM users WHERE serial_number = ?",
            (cert.serial_number,),
        ).fetchone()

        if existing:
            raise UserAlreadyExistsError(
                f"Certificate already registered to '{existing['username']}'."
            )

        cur = conn.execute(
            """INSERT INTO users(username, serial_number, subject, issuer, registered_at)
               VALUES (?, ?, ?, ?, ?)""",
            (username, cert.serial_number, cert.subject,
             cert.issuer, datetime.now().isoformat()),
        )
        user_id = cur.lastrowid

    _write_audit(
        "REGISTER",
        f"id={user_id} user={username} serial={cert.serial_number}",
    )
    _log.info("Registered user '%s' (serial=%s)", username, cert.serial_number)

    return RegisteredUser(
        id            = user_id,
        username      = username,
        serial_number = cert.serial_number,
        subject       = cert.subject,
        issuer        = cert.issuer,
        registered_at = datetime.now().isoformat(),
    )


def verify_signature(pdf_path: Path) -> VerificationResult:
    """Verify whether the signer of *pdf_path* is present in the registry.

    This function is deliberately *exception-free*: all errors are returned
    as a ``VerificationResult`` with ``status="ERROR"``.

    Parameters
    ----------
    pdf_path:
        Path to the signed PDF to check.

    Returns
    -------
    VerificationResult
        ``status`` is ``"APPROVED"``, ``"INVALID"``, or ``"ERROR"``.
    """
    try:
        cert = extract_signature_info(pdf_path)

        with _db() as conn:
            row = conn.execute(
                "SELECT username, serial_number FROM users WHERE serial_number = ?",
                (cert.serial_number,),
            ).fetchone()

        if row:
            _write_audit(
                "VERIFY_OK",
                f"pdf={pdf_path.name} user={row['username']}",
            )
            return VerificationResult(
                status        = "APPROVED",
                username      = row["username"],
                serial_number = row["serial_number"],
            )

        _write_audit(
            "VERIFY_FAIL",
            f"pdf={pdf_path.name} serial={cert.serial_number}",
        )
        return VerificationResult(
            status  = "INVALID",
            message = "Certificate not registered in this system.",
        )

    except (FileValidationError, CertificateError, DatabaseError) as exc:
        _log.error("Verification error for '%s': %s", pdf_path, exc)
        _write_audit("VERIFY_ERROR", f"pdf={pdf_path.name} error={exc}")
        return VerificationResult(status="ERROR", message=str(exc))

    except Exception as exc:
        _log.exception("Unexpected verification error for '%s'", pdf_path)
        _write_audit("VERIFY_ERROR", f"pdf={pdf_path.name} error={exc}")
        return VerificationResult(status="ERROR", message=f"Unexpected error: {exc}")


def list_users() -> List[RegisteredUser]:
    """Return all registered users ordered by registration date (oldest first)."""
    with _db() as conn:
        rows = conn.execute(
            "SELECT id, username, serial_number, subject, issuer, registered_at "
            "FROM users ORDER BY registered_at ASC"
        ).fetchall()
    return [_row_to_user(r) for r in rows]


def search_users(query: str) -> List[RegisteredUser]:
    """Return users whose *username* or *serial_number* contains *query*.

    The search is case-insensitive and supports partial matches.
    """
    if not query.strip():
        raise ValueError("Search query cannot be empty.")
    pattern = f"%{query.strip()}%"
    with _db() as conn:
        rows = conn.execute(
            """SELECT id, username, serial_number, subject, issuer, registered_at
               FROM users
               WHERE username      LIKE ? COLLATE NOCASE
                  OR serial_number LIKE ?""",
            (pattern, pattern),
        ).fetchall()
    return [_row_to_user(r) for r in rows]


def revoke_user(serial_number: str) -> bool:
    """Remove a registered user identified by their certificate *serial_number*.

    Returns
    -------
    bool
        ``True`` if a record was deleted, ``False`` if the serial was not found.
    """
    with _db() as conn:
        cur     = conn.execute(
            "DELETE FROM users WHERE serial_number = ?", (serial_number,)
        )
        deleted = cur.rowcount > 0

    if deleted:
        _write_audit("REVOKE", f"serial={serial_number}")
        _log.info("Revoked user with serial=%s", serial_number)
    return deleted


def export_users_csv() -> Path:
    """Export all users to a timestamped CSV file.

    The file is created inside the *exports/* directory (configurable via
    ``DSS_EXPORT_DIR``).

    Returns
    -------
    Path
        Absolute path to the generated CSV file.
    """
    CFG.export_dir.mkdir(parents=True, exist_ok=True)
    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = CFG.export_dir / f"users_{ts}.csv"
    users    = list_users()

    with out_path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(
            ["ID", "Username", "Serial Number", "Subject", "Issuer", "Registered At"]
        )
        for u in users:
            writer.writerow(
                [u.id, u.username, u.serial_number,
                 u.subject, u.issuer, u.registered_at]
            )

    _write_audit("EXPORT_CSV", f"path={out_path} count={len(users)}")
    _log.info("Exported %d user(s) to %s", len(users), out_path)
    return out_path.resolve()


# ══════════════════════════════════════════════════════════════════════════════
# §10  CLI PRESENTATION LAYER
# ══════════════════════════════════════════════════════════════════════════════

# Clamp terminal width between 60 and 90 characters for safe rendering.
_TW: int = max(60, min(shutil.get_terminal_size(fallback=(80, 24)).columns, 90))


# ── Layout primitives ─────────────────────────────────────────────────────────

def _banner(title: str) -> None:
    """Print a full-width cyan banner with *title*."""
    line = "═" * _TW
    print(_C.cyan(_C.bold(f"\n{line}")))
    print(_C.cyan(_C.bold(f"  {title}")))
    print(_C.cyan(_C.bold(line)))


def _section(title: str) -> None:
    """Print a section sub-header."""
    pad = max(0, _TW - len(title) - 8)
    print(_C.yellow(f"\n  ┌── {title} {'─' * pad}"))


def _kv(
    key: str,
    value: str,
    colour_fn: Optional[Callable[[str], str]] = None,
) -> None:
    """Print a labelled key-value row, optionally colouring the value."""
    # Pad on plain string BEFORE applying dim, so ANSI codes don't skew width.
    label = f"  {key}:".ljust(24)
    v     = colour_fn(value) if colour_fn else value
    print(f"{_C.dim(label)}{v}")


def _ok(msg: str)   -> None: print(_C.green(f"\n  ✔  {msg}"))
def _err(msg: str)  -> None: print(_C.red(f"\n  ✖  {msg}"))
def _warn(msg: str) -> None: print(_C.yellow(f"\n  ⚠  {msg}"))
def _pause()        -> None: input(_C.dim("\n  [Press Enter to continue]"))


# ── Input helpers ─────────────────────────────────────────────────────────────

def _prompt(label: str, default: str = "") -> str:
    """Display a labelled prompt and return the stripped input.

    Raises ``KeyboardInterrupt`` unchanged so the main loop can handle it.
    """
    hint = f" [{default}]" if default else ""
    val  = input(f"  {label}{hint}: ").strip()
    return val or default


def _get_pdf_path(label: str = "Path to signed PDF") -> Optional[Path]:
    """Prompt for a file path and validate it exists.

    Returns ``None`` (and prints an error) if the path is invalid.
    """
    raw = _prompt(label)
    if not raw:
        _err("No path provided.")
        return None
    p = Path(raw)
    if not p.is_file():
        _err(f"File not found: {p}")
        return None
    if p.suffix.lower() != ".pdf":
        _warn("File extension is not .pdf — proceeding anyway.")
    return p


# ── Table renderer ────────────────────────────────────────────────────────────

def _print_user_table(users: List[RegisteredUser]) -> None:
    """Render *users* as a fixed-width console table."""
    if not users:
        _warn("No users found.")
        return

    # Column widths: ID | Username | Serial Number | Registered At
    W   = [4, 22, 32, 19]
    HDR = ["ID", "Username", "Serial Number", "Registered At"]
    ROW = "  {:<{w0}}  {:<{w1}}  {:<{w2}}  {:<{w3}}"
    SEP = "  " + "  ".join("─" * w for w in W)

    print()
    print(_C.bold(ROW.format(*HDR, w0=W[0], w1=W[1], w2=W[2], w3=W[3])))
    print(_C.dim(SEP))

    for u in users:
        reg = u.registered_at[:19].replace("T", " ")
        print(ROW.format(
            str(u.id),
            u.username[:W[1]],
            u.serial_number[:W[2]],
            reg,
            w0=W[0], w1=W[1], w2=W[2], w3=W[3],
        ))

    print(_C.dim(f"\n  {len(users)} record(s)"))


# ── Detail view ───────────────────────────────────────────────────────────────

def _print_user_detail(u: RegisteredUser, heading: str = "Details") -> None:
    """Render full details for a single user."""
    _section(heading)
    _kv("ID",            str(u.id))
    _kv("Username",      u.username,       _C.green)
    _kv("Serial Number", u.serial_number)
    _kv("Subject",       u.subject)
    _kv("Issuer",        u.issuer)
    _kv("Registered At", u.registered_at[:19].replace("T", " "))


# ══════════════════════════════════════════════════════════════════════════════
# §11  MENU ACTIONS
# ══════════════════════════════════════════════════════════════════════════════

def _action_register() -> None:
    _banner("REGISTER NEW USER")

    username = _prompt("Username")
    if not username:
        _err("Username is required.")
        _pause()
        return

    pdf_path = _get_pdf_path()
    if pdf_path is None:
        _pause()
        return

    try:
        user = register_user(username, pdf_path)
        _ok("User registered successfully!")
        _print_user_detail(user, heading="Registration Details")

    except UserAlreadyExistsError as exc:
        _warn(str(exc))
    except (FileValidationError, CertificateError, ValueError) as exc:
        _err(str(exc))
    except DatabaseError as exc:
        _err(f"Database error: {exc}")

    _pause()


def _action_verify() -> None:
    _banner("VERIFY PDF SIGNATURE")

    pdf_path = _get_pdf_path()
    if pdf_path is None:
        _pause()
        return

    result = verify_signature(pdf_path)
    _section("Verification Result")

    if result.status == "APPROVED":
        _kv("Status",        result.status,             _C.green)
        _kv("Username",      result.username      or "", _C.green)
        _kv("Serial Number", result.serial_number or "")
        _ok("Signature APPROVED — signer is registered in this system.")

    elif result.status == "INVALID":
        _kv("Status",  result.status,         _C.red)
        _kv("Message", result.message or "", _C.red)
        _err("Signature INVALID — certificate not found in registry.")

    else:
        _kv("Status",  result.status,         _C.yellow)
        _kv("Message", result.message or "", _C.yellow)
        _err("Verification ERROR — see logs/ for details.")

    _pause()


def _action_list_users() -> None:
    _banner("ALL REGISTERED USERS")
    _print_user_table(list_users())
    _pause()


def _action_search() -> None:
    _banner("SEARCH USERS")

    query = _prompt("Search by username or serial number")
    if not query:
        _err("Search query cannot be empty.")
        _pause()
        return

    try:
        users = search_users(query)
    except ValueError as exc:
        _err(str(exc))
        _pause()
        return

    if users:
        _print_user_table(users)
    else:
        _warn(f"No results matching '{query}'.")

    _pause()


def _action_revoke() -> None:
    _banner("REVOKE USER")

    serial = _prompt("Certificate serial number to revoke")
    if not serial:
        _err("Serial number is required.")
        _pause()
        return

    # Safety confirmation — explicit "yes" required.
    confirm = _prompt(f"Type 'yes' to confirm permanent revocation")
    if confirm.lower() != "yes":
        _warn("Revocation cancelled.")
        _pause()
        return

    if revoke_user(serial):
        _ok(f"User with serial '{serial}' has been permanently revoked.")
    else:
        _warn(f"No user found with serial '{serial}'.")

    _pause()


def _action_export() -> None:
    _banner("EXPORT USERS TO CSV")

    try:
        path = export_users_csv()
        _ok("Export completed successfully.")
        _kv("Output file", str(path))
    except Exception as exc:
        _err(f"Export failed: {exc}")

    _pause()


# ══════════════════════════════════════════════════════════════════════════════
# §12  MAIN — interactive CLI loop
# ══════════════════════════════════════════════════════════════════════════════

# Menu definition: (key, label, handler)
# A handler of ``None`` signals the exit option.
_MENU: List[Tuple[str, str, Optional[Callable[[], None]]]] = [
    ("1", "Register User",        _action_register),
    ("2", "Verify PDF Signature", _action_verify),
    ("3", "List All Users",       _action_list_users),
    ("4", "Search Users",         _action_search),
    ("5", "Revoke User",          _action_revoke),
    ("6", "Export Users to CSV",  _action_export),
    ("0", "Exit",                 None),
]

_SENTINEL = object()   # Unique marker for "key not found in menu"


def _print_main_menu() -> None:
    """Render the application's main navigation menu."""
    _banner(f"{__app_name__}  •  v{__version__}")
    for key, label, fn in _MENU:
        bullet = _C.red(f"[{key}]") if key == "0" else _C.cyan(f"[{key}]")
        print(f"  {_C.bold(bullet)}  {label}")
    print()


def main() -> None:
    """Application entry point — runs the interactive CLI event loop."""
    init_database()
    _log.info("Application started — v%s", __version__)

    while True:
        try:
            _print_main_menu()
            choice = _prompt("Enter choice")

            # Resolve the handler for the chosen key.
            handler = next(
                (fn for k, _, fn in _MENU if k == choice),
                _SENTINEL,
            )

            if handler is _SENTINEL:
                _err("Invalid choice — please enter a number from the menu.")
                _pause()

            elif handler is None:
                # Exit option selected.
                print(_C.cyan("\n  Goodbye!\n"))
                _log.info("Application exited normally.")
                sys.exit(0)

            else:
                handler()   # type: ignore[operator]

        except KeyboardInterrupt:
            print(_C.yellow("\n\n  Interrupted — exiting."))
            _log.warning("Interrupted via KeyboardInterrupt.")
            sys.exit(0)

        except Exception as exc:
            _err(f"Unexpected error: {exc}")
            _log.exception("Unhandled exception in main event loop")
            _pause()


if __name__ == "__main__":
    main()