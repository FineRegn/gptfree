from __future__ import annotations

from typing import Any

from flask import jsonify

from outlook_web.repositories import overview as overview_repo
from outlook_web.security.auth import login_required


@login_required
def api_get_overview_summary() -> Any:
    return jsonify(overview_repo.get_overview_summary())


@login_required
def api_get_overview_verification() -> Any:
    return jsonify(overview_repo.get_verification_stats())


@login_required
def api_get_overview_external_api() -> Any:
    return jsonify(overview_repo.get_external_api_stats())


@login_required
def api_get_overview_pool() -> Any:
    return jsonify(overview_repo.get_pool_stats())


@login_required
def api_get_overview_activity() -> Any:
    return jsonify(overview_repo.get_activity_stats())
