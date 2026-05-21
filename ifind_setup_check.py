"""Minimal iFinD API setup check.

Credentials are read from environment variables:
    IFIND_USERNAME
    IFIND_PASSWORD
"""

import os
import sys
from typing import Any, Tuple


def require_credentials() -> Tuple[str, str]:
    username = os.environ.get("IFIND_USERNAME")
    password = os.environ.get("IFIND_PASSWORD")
    if not username or not password:
        raise RuntimeError(
            "Missing IFIND_USERNAME or IFIND_PASSWORD. Set them in the current "
            "PowerShell session before running this script."
        )
    return username, password


def import_ifind() -> Any:
    try:
        import iFinDPy as ifind
    except ImportError as exc:
        raise RuntimeError(
            "iFinDPy is not installed in this Python environment. "
            "Run: python -m pip install iFinDAPI"
        ) from exc
    return ifind


def ensure_ok(result: Any, action: str) -> None:
    error_code = getattr(result, "errorcode", result)
    if error_code != 0:
        raise RuntimeError(f"{action} failed, errorcode={error_code}, result={result!r}")


def query_usage_quota(ifind: Any) -> None:
    if hasattr(ifind, "THS_DataStatistics"):
        quota = ifind.THS_DataStatistics()
        print("Data statistics:")
        print(quota)
    else:
        print("THS_DataStatistics is unavailable in this iFinDPy version.")


def query_etf_sample(ifind: Any) -> None:
    """Small ETF sample: Shanghai Composite ETF short name."""
    result = ifind.THS_BD("510050.SH", "ths_fund_short_name_fund", "")
    print("ETF sample:")
    print(result)


def main() -> int:
    username, password = require_credentials()
    ifind = import_ifind()

    login_result = ifind.THS_iFinDLogin(username, password)
    ensure_ok(login_result, "iFinD login")
    print("iFinD login succeeded.")

    try:
        query_usage_quota(ifind)
        query_etf_sample(ifind)
    finally:
        if hasattr(ifind, "THS_iFinDLogout"):
            ifind.THS_iFinDLogout()
            print("iFinD logout completed.")

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
