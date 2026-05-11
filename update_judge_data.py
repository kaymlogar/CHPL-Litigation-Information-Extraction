#!/usr/bin/env python3
"""Download latest FJC judges export to local judges.csv."""

from __future__ import annotations

import os
import ssl
import sys
from pathlib import Path
from urllib.error import URLError
from urllib.request import Request, urlopen


FJC_JUDGES_CSV_URL = "https://www.fjc.gov/sites/default/files/history/judges.csv"
USER_AGENT = "Mozilla/5.0 (compatible; ONeillTrackerData/1.0)"
OUTPUT_FILENAME = "Federal Judicial Center Export.csv"

def _ssl_verify_disabled() -> bool:
    return os.environ.get("UPDATE_JUDGE_INSECURE_SSL", "").strip().lower() in (
        "1",
        "true",
        "yes",
    )


def _fetch_csv_bytes(url: str) -> bytes:
    """GET url with SSL verification.

    Attempt order:
    1. certifi CA bundle (fixes macOS Python cert issue without any user setup)
    2. Default system CA bundle
    3. Unverified (last resort, with warning — handles corporate proxies)
    """
    req = Request(url, headers={"User-Agent": USER_AGENT})
    timeout = 120

    if _ssl_verify_disabled():
        print(
            "Using unverified SSL (UPDATE_JUDGE_INSECURE_SSL is set).",
            file=sys.stderr,
        )
        ctx = ssl._create_unverified_context()
        with urlopen(req, timeout=timeout, context=ctx) as resp:
            return resp.read()

    # 1. Try certifi CA bundle — resolves the macOS Python cert issue silently.
    try:
        import certifi
        ctx = ssl.create_default_context(cafile=certifi.where())
        with urlopen(req, timeout=timeout, context=ctx) as resp:
            return resp.read()
    except ImportError:
        pass
    except URLError:
        pass

    # 2. Try default system CA bundle.
    try:
        with urlopen(req, timeout=timeout) as resp:
            return resp.read()
    except URLError as e:
        reason = e.reason
        if not isinstance(reason, ssl.SSLCertVerificationError):
            raise

    # 3. Last resort: unverified (e.g. corporate proxy with custom root cert).
    print(
        "Warning: SSL certificate verification failed. Retrying without verification.\n"
        "To fix: install the certifi package (`pip install certifi`) or set "
        "UPDATE_JUDGE_INSECURE_SSL=1 to suppress this message.",
        file=sys.stderr,
    )
    ctx = ssl._create_unverified_context()
    with urlopen(req, timeout=timeout, context=ctx) as resp:
        return resp.read()


def download_updated_judge_data() -> Path:
    """
    Download the FJC Biographical Directory export (organized by judge)
    and save it next to this script.
    """
    script_dir = Path(__file__).resolve().parent
    output_path = script_dir / OUTPUT_FILENAME
    data = _fetch_csv_bytes(FJC_JUDGES_CSV_URL)
    output_path.write_bytes(data)
    return output_path


def main() -> None:
    
    output_path = download_updated_judge_data()
    print(f"Downloaded updated judge data to: {output_path}")


if __name__ == "__main__":
    main()
