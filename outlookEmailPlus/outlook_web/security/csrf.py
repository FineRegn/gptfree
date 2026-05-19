from __future__ import annotations

from typing import Callable, Optional, Tuple

try:
    from flask_wtf.csrf import CSRFProtect
    from flask_wtf.csrf import generate_csrf as _generate_csrf  # type: ignore

    CSRF_AVAILABLE = True
except ImportError:  # pragma: no cover
    CSRFProtect = None  # type: ignore
    _generate_csrf = None  # type: ignore
    CSRF_AVAILABLE = False
    print("Warning: flask-wtf not installed. CSRF protection is disabled. Install with: pip install flask-wtf")


def generate_csrf():
    """生成 CSRF token（模块级别导出）"""
    if _generate_csrf is not None:
        return _generate_csrf()
    return None


def init_csrf(app) -> Tuple[Optional[object], Callable, Optional[Callable]]:
    """
    初始化 CSRF 保护（如果可用）
    返回：csrf 实例、csrf_exempt 装饰器、generate_csrf 函数（可能为 None）
    """
    if CSRF_AVAILABLE:
        csrf = CSRFProtect(app)  # type: ignore[misc]
        app.config["WTF_CSRF_TIME_LIMIT"] = None  # CSRF token 不过期
        app.config["WTF_CSRF_SSL_STRICT"] = False  # 允许非HTTPS环境（开发环境）
        print("CSRF protection enabled")

        def csrf_exempt(f):
            return csrf.exempt(f)

        return csrf, csrf_exempt, _generate_csrf

    # 显式禁用 CSRF 保护
    app.config["WTF_CSRF_ENABLED"] = False
    app.config["WTF_CSRF_CHECK_DEFAULT"] = False
    print("CSRF protection disabled")

    def csrf_exempt(f):
        return f

    return None, csrf_exempt, None
