"""过期清理与录入回归测试：仅使用临时 SQLite，绝不连接线上 Turso。"""
import os
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import pytest
import streamlit as st

_TEST_DATA_DIR = tempfile.mkdtemp(prefix="lecture-cleanup-tests-")
os.environ["DATA_DIR"] = _TEST_DATA_DIR
# load_dotenv 默认不覆盖已有环境变量，因此空值可确保测试不会使用 .env 中的 Turso。
os.environ["TURSO_DB_URL"] = ""
os.environ["TURSO_AUTH_TOKEN"] = ""
# 同时屏蔽用户级 Streamlit Secrets，形成本地测试的双重隔离。
st.secrets = {}

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import app  # noqa: E402

NOW = datetime(2026, 10, 28, 19, 0, 0)


def _insert_raw(title, category, time_start=None, time_end=None):
    conn = app.get_conn()
    try:
        cur = conn.execute(
            "INSERT INTO lectures (title, category, time_start, time_end) VALUES (?, ?, ?, ?)",
            (title, category, time_start, time_end),
        )
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def _titles():
    conn = app.get_conn()
    try:
        return {row[0] for row in conn.execute("SELECT title FROM lectures").fetchall()}
    finally:
        conn.close()


def _record(title="新讲座", time_start="2026-10-29 10:00:00", time_end="2026-10-29 11:00:00"):
    return {
        "title": title,
        "category": "lecture",
        "time_start": time_start,
        "time_end": time_end,
    }


@pytest.fixture(autouse=True)
def empty_database():
    conn = app.get_conn()
    try:
        conn.execute("DELETE FROM lectures")
        conn.commit()
    finally:
        conn.close()
    yield


def test_cleanup_handles_formats_categories_and_invalid_times():
    _insert_raw("过期讲座", "lecture", "2026-10-20 10:00:00", "2026-10-20 11:00:00")
    _insert_raw("T格式过期活动", "humanities_activity", "2026-10-28T10:00:00", "2026-10-28T11:00:00")
    _insert_raw("未过期活动", "humanities_activity", "2026-10-29T10:00:00", "2026-10-29T11:00:00")
    _insert_raw("场所永久保留", "study_venue", "2020-01-01 08:00:00", "2020-01-01 18:00:00")
    _insert_raw("无时间保留", "lecture", None, None)
    _insert_raw("坏时间保留", "lecture", "not-a-date", "invalid")

    assert app.cleanup_expired(NOW) == 2
    assert _titles() == {"未过期活动", "场所永久保留", "无时间保留", "坏时间保留"}


def test_cleanup_uses_end_then_start_and_exact_boundary_is_expired():
    _insert_raw("开始早但结束晚", "lecture", "2026-10-20 10:00:00", "2026-10-29 11:00:00")
    _insert_raw("无结束用开始", "lecture", "2026-10-20 10:00:00", None)
    _insert_raw("刚好到结束", "humanities_activity", "2026-10-28 18:00:00", "2026-10-28 19:00:00")

    assert app.cleanup_expired(NOW) == 2
    assert _titles() == {"开始早但结束晚"}


def test_cleanup_normalizes_timezone_and_current_time_zone():
    _insert_raw("UTC时间到界", "lecture", "2026-10-28T10:00:00+00:00", "2026-10-28T11:00:00+00:00")
    aware_now = datetime(2026, 10, 28, 11, 0, 0, tzinfo=app.timezone.utc)

    assert app.cleanup_expired(aware_now) == 1
    assert _titles() == set()


def test_successful_insert_cleans_expired_and_preserves_valid(monkeypatch):
    _insert_raw("待清理", "humanities_activity", "2026-10-28T10:00:00", "2026-10-28T11:00:00")
    _insert_raw("原有有效", "lecture", "2026-10-30 10:00:00", "2026-10-30 11:00:00")
    real_cleanup = app.cleanup_expired
    monkeypatch.setattr(app, "cleanup_expired", lambda: real_cleanup(NOW))

    ok, message = app.insert_lecture(_record())

    assert ok is True
    assert "清理 1 条" in message
    assert _titles() == {"原有有效", "新讲座"}


def test_duplicate_insert_still_cleans_without_adding_record(monkeypatch):
    _insert_raw("重复讲座", "lecture", "2026-10-29 10:00:00", "2026-10-29 11:00:00")
    _insert_raw("重复时清理", "lecture", "2026-10-20 10:00:00", "2026-10-20 11:00:00")
    real_cleanup = app.cleanup_expired
    monkeypatch.setattr(app, "cleanup_expired", lambda: real_cleanup(NOW))

    ok, message = app.insert_lecture(_record(title=" 重复，讲座 "))

    assert ok is False
    assert "重复" in message
    assert "清理 1 条" in message
    assert _titles() == {"重复讲座"}


def test_unique_constraint_skip_still_cleans(monkeypatch):
    _insert_raw("唯一约束重复", "lecture", "2026-10-29 10:00:00", "2026-10-29 11:00:00")
    _insert_raw("约束跳过时清理", "lecture", "2026-10-20 10:00:00", "2026-10-20 11:00:00")
    real_cleanup = app.cleanup_expired
    monkeypatch.setattr(app, "cleanup_expired", lambda: real_cleanup(NOW))
    # 绕过模糊查重，专门验证数据库 UNIQUE(title, time_start) 的跳过路径。
    monkeypatch.setattr(app, "is_duplicate", lambda record: False)

    ok, message = app.insert_lecture(_record(title="唯一约束重复"))

    assert ok is False
    assert "重复" in message
    assert "清理 1 条" in message
    assert _titles() == {"唯一约束重复"}


def test_new_expired_record_is_removed_immediately(monkeypatch):
    real_cleanup = app.cleanup_expired
    monkeypatch.setattr(app, "cleanup_expired", lambda: real_cleanup(NOW))

    ok, message = app.insert_lecture(
        _record(title="误录过期", time_start="2026-10-28T10:00:00", time_end="2026-10-28T11:00:00")
    )

    assert ok is True
    assert "清理 1 条" in message
    assert _titles() == set()
