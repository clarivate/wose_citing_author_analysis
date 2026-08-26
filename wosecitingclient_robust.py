from __future__ import annotations

import time
import random
from typing import Any, List, Dict, Optional

import requests


class WoSAuthenticationError(Exception):
    """Raised when the WoS API returns 401/403 (invalid key or insufficient access)."""
    pass


BASE_URL = "https://wos-api.clarivate.com/api/wos/citing"

CONNECT_TIMEOUT = 10   # seconds
READ_TIMEOUT = 60      # seconds
PAGE_THROTTLE = 0.25   # seconds


def _request_with_retries(
    url: str,
    headers: Dict[str, str],
    params: Dict[str, Any],
    max_tries: int = 6,
) -> requests.Response:
    """
    Make a GET request with retry handling for transient WoS API errors.
    """

    backoff = 0.5

    for attempt in range(1, max_tries + 1):
        try:
            r = requests.get(
                url,
                headers=headers,
                params=params,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )

            # Retry transient API/server/rate-limit errors
            if r.status_code in (429, 500, 502, 503, 504):
                retry_after = r.headers.get("Retry-After")

                try:
                    wait = float(retry_after) if retry_after else None
                except (TypeError, ValueError):
                    wait = None

                if wait is None:
                    wait = backoff + random.random() * 0.5

                time.sleep(wait)
                backoff = min(backoff * 2, 8.0)
                continue

            # Friendly message for invalid/unauthorized API key
            if r.status_code in (401, 403):
                raise WoSAuthenticationError(
                    f"\nWeb of Science API authentication/authorization failed "
                    f"(HTTP {r.status_code}).\n\n"
                    f"This usually means your API key is invalid/expired, or "
                    f"the key does not have access to this endpoint/collection.\n\n"
                    f"Check that EXPANDED_APIKEY (or the key you passed) is "
                    f"correct, active, and has WoS Expanded API access.\n"
                )

            r.raise_for_status()
            return r

        except requests.exceptions.RequestException:
            if attempt == max_tries:
                raise

            time.sleep(backoff + random.random() * 0.5)
            backoff = min(backoff * 2, 8.0)

    raise RuntimeError("Unreachable")


def _normalize_data(data: Any) -> List[Dict[str, Any]]:
    """
    Normalize the WoS 'Data' element into a list.

    Normally Data is a list, but this protects downstream code if
    an empty value or individual object is returned.
    """

    if not data:
        return []

    if isinstance(data, list):
        return data

    if isinstance(data, dict):
        return [data]

    return []


def get_response(
    apikey: str,
    params: Dict[str, Any],
    firstRecord: int = 1,
    count: int = 50,
    url: Optional[str] = None,
) -> Dict[str, Any] | None:
    """
    Retrieve one page from the WoS citing endpoint.
    """

    headers = {
        "Accept": "application/json",
        "X-ApiKey": apikey,
    }

    params = dict(params)
    params["count"] = count
    params["firstRecord"] = firstRecord

    req_url = BASE_URL + url if url else BASE_URL

    r = _request_with_retries(
        req_url,
        headers,
        params,
    )

    time.sleep(PAGE_THROTTLE)

    return r.json()


def get_addl_results(
    apikey: str,
    params: Dict[str, Any],
    recordsFound: int,
    firstRecord: int = 1,
    count: int = 50,
    data: Optional[List[Dict[str, Any]]] = None,
) -> List[Dict[str, Any]]:
    """
    Retrieve remaining pages of citing records.
    """

    headers = {
        "Accept": "application/json",
        "X-ApiKey": apikey,
    }

    if data is None:
        data = []

    params = dict(params)
    params["count"] = count

    while firstRecord <= recordsFound:
        params["firstRecord"] = firstRecord

        r = _request_with_retries(
            BASE_URL,
            headers,
            params,
        )

        js = r.json()

        page_data = _normalize_data(js.get("Data", []))

        if page_data:
            data.extend(page_data)

        firstRecord += count

        time.sleep(PAGE_THROTTLE)

    return data


def get_all_records(
    apikey: str,
    params: Dict[str, Any],
    firstRecord: int = 1,
    count: int = 50,
) -> List[Dict[str, Any]]:
    """
    Retrieve all citing records for the supplied parameters.
    """

    r = get_response(
        apikey,
        params,
        firstRecord,
        count,
    )

    if r is None:
        return []

    qr = r.get("QueryResult", {})

    recordsFound = qr.get("RecordsFound", 0) or 0

    data = _normalize_data(
        r.get("Data", [])
    )

    if recordsFound > len(data):
        data = get_addl_results(
            apikey,
            params,
            recordsFound,
            firstRecord + count,
            count,
            data,
        )

    return data


def get_citing_records(
    apikey: str,
    unique_id: str,
    database_id: str = "WOS",
    count: int = 50,
) -> List[Dict[str, Any]]:
    """
    Convenience function to retrieve all citing records for a single WoS UID.

    Example:
        get_citing_records(
            apikey,
            "WOS:000486006400008"
        )
    """

    params = {
        "databaseId": database_id,
        "uniqueId": unique_id,
    }

    return get_all_records(
        apikey,
        params,
        firstRecord=1,
        count=count,
    )