#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""一键释放 OutlookEmailPlus 邮箱池中被领取但未完成的账号。"""

from __future__ import annotations

import argparse
import sqlite3
import sys
from datetime import datetime, timezone
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parent
DEFAULT_DB_PATH = ROOT_DIR / "outlookEmailPlus" / "data" / "outlook_accounts.db"
DEFAULT_REASON = "admin force release by release_outlook_email_pool.py"


def utc_now_text() -> str:
    return datetime.now(timezone.utc).replace(tzinfo=None).isoformat() + "Z"


def parse_claimed_by(value: str | None) -> tuple[str, str]:
    if not value:
        return "", ""
    caller_id, separator, task_id = value.partition(":")
    if not separator:
        return caller_id, ""
    return caller_id, task_id


def connect_database(db_path: Path) -> sqlite3.Connection:
    if not db_path.exists():
        raise FileNotFoundError(f"数据库不存在: {db_path}")

    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    return conn


def load_claimed_accounts(conn: sqlite3.Connection) -> list[sqlite3.Row]:
    return conn.execute(
        """
        SELECT id, email, provider, account_type, claimed_by, claimed_at,
               lease_expires_at, claim_token, claimed_project_key
        FROM accounts
        WHERE pool_status = 'claimed'
        ORDER BY id
        """
    ).fetchall()


def print_accounts(accounts: list[sqlite3.Row]) -> None:
    if not accounts:
        print("没有处于 claimed 状态的邮箱，无需释放。")
        return

    print(f"待释放邮箱数量: {len(accounts)}")
    for account in accounts:
        print(
            "- "
            f"id={account['id']} "
            f"email={account['email']} "
            f"provider={account['provider'] or ''} "
            f"claimed_by={account['claimed_by'] or ''} "
            f"lease_expires_at={account['lease_expires_at'] or ''}"
        )


def release_accounts(conn: sqlite3.Connection, accounts: list[sqlite3.Row], reason: str) -> int:
    if not accounts:
        return 0

    now_text = utc_now_text()
    try:
        conn.execute("BEGIN IMMEDIATE")
        released_count = 0
        for account in accounts:
            caller_id, task_id = parse_claimed_by(account["claimed_by"])
            cursor = conn.execute(
                """
                UPDATE accounts SET
                    pool_status = 'available',
                    claimed_by = NULL,
                    claimed_at = NULL,
                    lease_expires_at = NULL,
                    claim_token = NULL,
                    claimed_project_key = NULL,
                    updated_at = ?
                WHERE id = ? AND pool_status = 'claimed'
                """,
                (now_text, account["id"]),
            )
            if cursor.rowcount != 1:
                continue

            conn.execute(
                """
                INSERT INTO account_claim_logs
                    (account_id, claim_token, caller_id, task_id, action, result, detail, created_at)
                VALUES (?, ?, ?, ?, 'release', 'manual_release', ?, ?)
                """,
                (
                    account["id"],
                    account["claim_token"] or "",
                    caller_id,
                    task_id,
                    reason,
                    now_text,
                ),
            )
            released_count += 1
        conn.execute("COMMIT")
    except Exception:
        conn.execute("ROLLBACK")
        raise

    return released_count


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="一键释放 OutlookEmailPlus 邮箱池 claimed 账号")
    parser.add_argument(
        "--db",
        default=str(DEFAULT_DB_PATH),
        help=f"SQLite 数据库路径，默认: {DEFAULT_DB_PATH}",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只打印将释放的邮箱，不修改数据库",
    )
    parser.add_argument(
        "--reason",
        default=DEFAULT_REASON,
        help="写入 account_claim_logs.detail 的释放原因",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    db_path = Path(args.db).expanduser().resolve()

    try:
        with connect_database(db_path) as conn:
            accounts = load_claimed_accounts(conn)
            print(f"数据库: {db_path}")
            print_accounts(accounts)

            if args.dry_run:
                print("dry-run 模式：未修改数据库。")
                return 0

            released_count = release_accounts(conn, accounts, args.reason)
            print(f"释放完成: {released_count} 个邮箱已恢复为 available。")
            return 0
    except Exception as exc:
        print(f"释放失败: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
