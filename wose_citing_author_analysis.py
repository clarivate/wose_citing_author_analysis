"""
Web of Science Citing Author Analysis

Retrieves a Web of Science source set as Short Records, identifies authors on
those records, retrieves citing articles as Short Records, and analyzes citing
authors using citation-link-weighted counts.

Required companion files:
- wosesrclient_robust.py
- wosecitingclient_robust.py

Environment:
EXPANDED_APIKEY=<your Web of Science Expanded API key>

Required packages:
requests
python-dotenv
pandas
xlsxwriter (recommended)
openpyxl (fallback)
"""

from __future__ import annotations

from collections import Counter, defaultdict
from typing import Any, Dict, List, Set, Tuple
from dotenv import load_dotenv

import argparse
import datetime
import os
import re
import sys
import unicodedata

import pandas as pd

import wosesrclient_robust
from wosesrclient_robust import InvalidWoSQueryError
from wosesrclient_robust import WoSAuthenticationError as SRAuthError

from wosecitingclient_robust import get_response as get_citing_response
from wosecitingclient_robust import WoSAuthenticationError as CitingAuthError

# =============================================================================
# Parameters
# =============================================================================

params = {
    "databaseId": "WOS",
    "usrQuery": "AU=(Stanwood)",  # Edit this query when running without -q/--query
    "firstRecord": 1,
    "count": 50,
    "optionView": "SR",
}

# r_id is the preferred grouping key.
# If an author has no r_id, True keeps that author using a normalized-name
# fallback rather than dropping the author entirely.
INCLUDE_AUTHORS_WITHOUT_RID = True

# =============================================================================
# CLI
# =============================================================================

def parse_args(args):
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "-q",
        "--query",
        help=(
            "Optional Web of Science query. If supplied, this overrides "
            "params['usrQuery'] in the script."
        ),
    )

    parser.add_argument(
        "-k",
        "--key",
        help=(
            "WoS Expanded API key. If omitted, EXPANDED_APIKEY is read "
            "from the .env file."
        ),
    )

    parser.add_argument(
        "-o",
        "--output",
        help="Optional Excel output filename.",
    )

    parser.add_argument(
        "--exclude-authors-without-rid",
        action="store_true",
        help=(
            "Exclude authors that do not have a Web of Science ResearcherID. "
            "Default behavior keeps them using a name-based fallback key."
        ),
    )

    return parser.parse_args(args)

# =============================================================================
# General helpers
# =============================================================================

def ensure_list(value: Any) -> List[Any]:
    if value is None:
        return []
    return value if isinstance(value, list) else [value]

def parse_uid(rec: Dict[str, Any]) -> str:
    if not isinstance(rec, dict):
        return ""
    return str(rec.get("UID") or "").strip()

def normalize_wos_uid(uid: Any) -> str:
    """Normalize Web of Science UIDs to a consistent WOS:-prefixed form."""
    value = str(uid or "").strip()
    if not value:
        return ""

    if value.upper().startswith("WOS:"):
        return "WOS:" + value.split(":", 1)[1]

    if re.fullmatch(r"\d+", value):
        return f"WOS:{value}"

    return value

def normalize_name(value: Any) -> str:
    return " ".join(str(value or "").strip().casefold().split())

def safe_slug_from_query(query: str) -> str:
    clean = re.sub(r"[^a-zA-Z0-9 ]", "", str(query or ""))
    return clean.strip().replace(" ", "_")[:40] or "query"

def choose_display_name(name_counter: Counter) -> str:
    """
    Use the most frequently observed preferred-name label.
    Resolve ties alphabetically for reproducible output.
    """
    if not name_counter:
        return ""

    top_count = max(name_counter.values())
    candidates = [
        name for name, count in name_counter.items()
        if count == top_count
    ]
    return sorted(candidates, key=lambda x: x.casefold())[0]

def progress(message: str = "") -> None:
    """Print an immediately flushed progress message."""
    print(message, flush=True)

def stage_banner(stage_number: int, title: str, total_stages: int = 6) -> None:
    print()
    progress(f"STAGE {stage_number}/{total_stages}: {title}")

def normalize_author_match_name(value: Any) -> str:
    """Normalize an author name for diagnostic source-set name matching."""
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(
        ch for ch in text
        if not unicodedata.combining(ch)
    )
    text = text.casefold()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())

# =============================================================================
# Author extraction
# =============================================================================

def extract_author_identity(
    name: Dict[str, Any],
    include_authors_without_rid: bool,
) -> Tuple[str, str, str, str]:
    """
    Return:
        group_key
        r_id
        display_label
        grouping_basis

    Preferred grouping:
        r_id

    Preferred label:
        preferred_name.full_name

    Display fallbacks:
        full_name
        display_name
        wos_standard

    Corporate/group authors are excluded from this person-author analysis.
    """
    if not isinstance(name, dict):
        return "", "", "", ""

    role = str(name.get("role") or "").strip().casefold()

    if role == "corp":
        return "", "", "", ""

    # Keep blank roles as a fallback, but reject non-author roles.
    if role and role != "author":
        return "", "", "", ""

    r_id = str(name.get("r_id") or "").strip()

    preferred_name = name.get("preferred_name")
    preferred_full_name = ""

    if isinstance(preferred_name, dict):
        preferred_full_name = str(
            preferred_name.get("full_name") or ""
        ).strip()

    display_label = (
        preferred_full_name
        or str(name.get("full_name") or "").strip()
        or str(name.get("display_name") or "").strip()
        or str(name.get("wos_standard") or "").strip()
    )

    if r_id:
        return f"RID:{r_id.upper()}", r_id, display_label, "ResearcherID"

    if not include_authors_without_rid:
        return "", "", "", ""

    normalized_label = normalize_name(display_label)
    if not normalized_label:
        return "", "", "", ""

    return (
        f"NAME:{normalized_label}",
        "",
        display_label,
        "Name fallback",
    )

def extract_authors_from_record(
    rec: Dict[str, Any],
    include_authors_without_rid: bool,
) -> List[Tuple[str, str, str, str]]:
    """
    Extract person authors from:
        static_data -> summary -> names -> name
    """
    try:
        names_block = (
            rec.get("static_data", {})
               .get("summary", {})
               .get("names", {})
        )
    except AttributeError:
        return []

    if not isinstance(names_block, dict):
        return []

    names = ensure_list(names_block.get("name"))

    authors = []

    for name in names:
        identity = extract_author_identity(
            name,
            include_authors_without_rid=include_authors_without_rid,
        )

        if identity[0]:
            authors.append(identity)

    return authors

# =============================================================================
# Citing endpoint normalization
# =============================================================================

def extract_citing_items(data: Any) -> List[Dict[str, Any]]:
    """
    Normalize Short Record items returned by the Expanded API /citing endpoint.

    Only dictionaries containing a UID are returned.
    """
    output: List[Dict[str, Any]] = []

    def visit(value: Any) -> None:
        if isinstance(value, list):
            for item in value:
                visit(item)
            return

        if not isinstance(value, dict):
            return

        if value.get("UID"):
            output.append(value)
            return

        # Common Expanded API response containers.
        if "Data" in value:
            visit(value.get("Data"))

        try:
            visit(
                value.get("Records", {})
                     .get("records", {})
                     .get("REC")
            )
        except AttributeError:
            pass

    visit(data)

    return output

# =============================================================================
# Source-author analysis
# =============================================================================

def build_source_author_analysis(
    source_records: List[Dict[str, Any]],
    include_authors_without_rid: bool,
) -> pd.DataFrame:
    """Summarize authors appearing on the source records."""
    stats = defaultdict(
        lambda: {
            "r_id": "",
            "names": Counter(),
            "grouping_basis": "",
            "paper_uids": set(),
        }
    )

    for rec in source_records:
        source_uid = parse_uid(rec)

        if not source_uid:
            continue

        seen_on_record: Set[str] = set()

        for group_key, r_id, display_label, grouping_basis in extract_authors_from_record(
            rec,
            include_authors_without_rid=include_authors_without_rid,
        ):
            if group_key in seen_on_record:
                continue

            seen_on_record.add(group_key)

            entry = stats[group_key]

            if r_id and not entry["r_id"]:
                entry["r_id"] = r_id

            if display_label:
                entry["names"][display_label] += 1

            if not entry["grouping_basis"]:
                entry["grouping_basis"] = grouping_basis

            entry["paper_uids"].add(source_uid)

    source_record_count = len({
        parse_uid(rec)
        for rec in source_records
        if parse_uid(rec)
    })

    rows = []

    for group_key, entry in stats.items():
        paper_count = len(entry["paper_uids"])

        rows.append(
            {
                "ResearcherID": entry["r_id"],
                "Preferred Name": choose_display_name(entry["names"]),
                "Paper Count": paper_count,
                "Percent of Source Papers": (
                    paper_count / source_record_count
                    if source_record_count
                    else 0.0
                ),
                "Grouping Basis": entry["grouping_basis"],
            }
        )

    columns = [
        "ResearcherID",
        "Preferred Name",
        "Paper Count",
        "Percent of Source Papers",
        "Grouping Basis",
    ]

    if not rows:
        return pd.DataFrame(columns=columns)

    return (
        pd.DataFrame(rows, columns=columns)
        .sort_values(
            ["Paper Count", "Preferred Name", "ResearcherID"],
            ascending=[False, True, True],
        )
        .reset_index(drop=True)
    )

# =============================================================================
# Source-set author matching reference
# =============================================================================

def build_source_author_reference(
    source_records: List[Dict[str, Any]],
    include_authors_without_rid: bool,
) -> Dict[str, Any]:
    """
    Build lookup structures used to compare citing authors with authors who
    occur in the original source set.

    Exact match:
        Same ResearcherID.

    Possible match:
        ResearcherID differs or is absent, but the cautiously normalized
        display/preferred name matches a source-set author name.

    Name matching never merges author groups; it only adds a diagnostic label.
    """
    source_rids: Set[str] = set()
    name_to_rids: Dict[str, Set[str]] = defaultdict(set)
    name_to_labels: Dict[str, Counter] = defaultdict(Counter)

    for rec in source_records:
        for _, r_id, display_label, _ in extract_authors_from_record(
            rec,
            include_authors_without_rid=include_authors_without_rid,
        ):
            rid_key = r_id.strip().upper() if r_id else ""
            normalized_name = normalize_author_match_name(display_label)

            if rid_key:
                source_rids.add(rid_key)

            if normalized_name:
                if rid_key:
                    name_to_rids[normalized_name].add(rid_key)
                if display_label:
                    name_to_labels[normalized_name][display_label] += 1

    return {
        "source_rids": source_rids,
        "name_to_rids": name_to_rids,
        "name_to_labels": name_to_labels,
    }

def classify_source_author_match(
    r_id: str,
    display_label: str,
    source_author_reference: Dict[str, Any],
) -> Tuple[str, str]:
    """
    Return:
        match_status
        matched_source_researcher_ids

    Status values:
        Exact Source Author
        Possible Source Author
        Not Matched to Source Set
    """
    rid_key = str(r_id or "").strip().upper()
    source_rids = source_author_reference.get("source_rids", set())

    if rid_key and rid_key in source_rids:
        return "Exact Source Author", r_id

    normalized_name = normalize_author_match_name(display_label)
    name_to_rids = source_author_reference.get("name_to_rids", {})
    name_to_labels = source_author_reference.get("name_to_labels", {})

    if normalized_name and normalized_name in name_to_labels:
        matched_rids = sorted(name_to_rids.get(normalized_name, set()))
        return (
            "Possible Source Author",
            "; ".join(matched_rids),
        )

    return "Not Matched to Source Set", ""

# =============================================================================
# Citation-link retrieval
# =============================================================================

def retrieve_citation_links(
    apikey: str,
    source_uids: List[str],
) -> Tuple[
    List[Tuple[str, str]],
    Dict[str, Dict[str, Any]],
    List[str],
]:
    """
    Retrieve citing articles as Short Records for every source UT.

    Each source-to-citing relationship is preserved as a citation link, while
    each unique citing Short Record is cached once for later author analysis.
    """
    citation_links: List[Tuple[str, str]] = []
    citing_records_by_uid: Dict[str, Dict[str, Any]] = {}
    citing_errors: List[str] = []

    total_sources = len(source_uids)
    page_size = 50

    progress(f"Processing {total_sources} source record(s) one at a time.")

    for source_index, source_uid in enumerate(source_uids, start=1):
        progress(
            f"Source {source_index}/{total_sources}: "
            f"retrieving citing Short Records for {source_uid}..."
        )

        citing_params = {
            "databaseId": "WOS",
            "uniqueId": source_uid,
            "optionView": "SR",
        }

        seen_for_source: Set[str] = set()
        source_citing_uids: List[str] = []

        try:
            first_response = get_citing_response(
                apikey,
                citing_params,
                firstRecord=1,
                count=page_size,
            )

            if first_response is None:
                records_found = 0
                first_items: List[Dict[str, Any]] = []
            else:
                query_result = first_response.get("QueryResult", {})
                records_found = int(
                    query_result.get("RecordsFound", 0) or 0
                )
                first_items = extract_citing_items(
                    first_response.get("Data", [])
                )

            total_pages = max(
                1,
                (records_found + page_size - 1) // page_size,
            )

            def add_items(items: List[Dict[str, Any]]) -> None:
                for item in items:
                    citing_uid = normalize_wos_uid(parse_uid(item))

                    if not citing_uid:
                        continue

                    if citing_uid not in citing_records_by_uid:
                        citing_records_by_uid[citing_uid] = item

                    if citing_uid in seen_for_source:
                        continue

                    seen_for_source.add(citing_uid)
                    source_citing_uids.append(citing_uid)

            add_items(first_items)

            if records_found:
                progress(
                    f"  Source {source_index}/{total_sources}: "
                    f"{records_found} citing article(s); "
                    f"page 1/{total_pages} retrieved "
                    f"({min(len(first_items), records_found)}/{records_found})."
                )
            else:
                progress(
                    f"  Source {source_index}/{total_sources}: "
                    "0 citing articles."
                )

            page_number = 1
            first_record = 1 + page_size

            while first_record <= records_found:
                page_number += 1

                page_response = get_citing_response(
                    apikey,
                    citing_params,
                    firstRecord=first_record,
                    count=page_size,
                )

                page_items = extract_citing_items(
                    (page_response or {}).get("Data", [])
                )
                add_items(page_items)

                retrieved_through = min(
                    first_record + page_size - 1,
                    records_found,
                )

                progress(
                    f"  Source {source_index}/{total_sources}: "
                    f"page {page_number}/{total_pages} retrieved "
                    f"({retrieved_through}/{records_found})."
                )

                first_record += page_size

        except CitingAuthError:
            raise

        except Exception as exc:
            message = f"{source_uid}: {exc}"
            citing_errors.append(message)
            progress(f"  Citing retrieval error: {exc}")
            continue

        for citing_uid in source_citing_uids:
            citation_links.append((source_uid, citing_uid))

        progress(
            f"  Completed source {source_index}/{total_sources}: "
            f"{len(source_citing_uids)} citing article(s); "
            f"{len(citation_links)} citation link(s) accumulated; "
            f"{len(citing_records_by_uid)} unique citing Short Record(s) cached."
        )

    return citation_links, citing_records_by_uid, citing_errors

# =============================================================================
# Citing-author citation contribution analysis
# =============================================================================

def build_citing_author_analysis(
    citing_records_by_uid: Dict[str, Dict[str, Any]],
    citation_links: List[Tuple[str, str]],
    source_records: List[Dict[str, Any]],
    include_authors_without_rid: bool,
) -> pd.DataFrame:
    """
    Build citation-link-weighted citing-author metrics.

    Citation Contribution counts source-to-citing citation links associated
    with each author. Source-set author matching uses ResearcherID first and
    normalized-name matching only as a diagnostic fallback.
    """
    link_count_by_citing_uid = Counter(
        citing_uid
        for _, citing_uid in citation_links
    )

    source_uids_by_citing_uid = defaultdict(set)

    for source_uid, citing_uid in citation_links:
        source_uids_by_citing_uid[citing_uid].add(source_uid)

    total_citation_links = len(citation_links)
    total_source_papers = len({
        normalize_wos_uid(parse_uid(rec))
        for rec in source_records
        if normalize_wos_uid(parse_uid(rec))
    })

    source_author_reference = build_source_author_reference(
        source_records,
        include_authors_without_rid=include_authors_without_rid,
    )

    stats = defaultdict(
        lambda: {
            "r_id": "",
            "names": Counter(),
            "grouping_basis": "",
            "citation_contribution": 0,
            "citing_uids": set(),
            "source_uids": set(),
        }
    )

    for citing_uid, rec in citing_records_by_uid.items():
        link_weight = int(link_count_by_citing_uid.get(citing_uid, 0))

        if link_weight <= 0:
            continue

        seen_on_record: Set[str] = set()

        for group_key, r_id, display_label, grouping_basis in extract_authors_from_record(
            rec,
            include_authors_without_rid=include_authors_without_rid,
        ):
            if group_key in seen_on_record:
                continue

            seen_on_record.add(group_key)
            entry = stats[group_key]

            if r_id and not entry["r_id"]:
                entry["r_id"] = r_id

            if display_label:
                entry["names"][display_label] += 1

            if not entry["grouping_basis"]:
                entry["grouping_basis"] = grouping_basis

            entry["citation_contribution"] += link_weight
            entry["citing_uids"].add(citing_uid)
            entry["source_uids"].update(
                source_uids_by_citing_uid.get(citing_uid, set())
            )

    rows = []

    for group_key, entry in stats.items():
        citation_contribution = int(entry["citation_contribution"])
        distinct_citing_articles = len(entry["citing_uids"])
        distinct_source_papers = len(entry["source_uids"])
        preferred_name = choose_display_name(entry["names"])

        match_status, matched_source_rids = classify_source_author_match(
            entry["r_id"],
            preferred_name,
            source_author_reference,
        )

        rows.append(
            {
                "ResearcherID": entry["r_id"],
                "Preferred Name": preferred_name,
                "Source Set Author Match": match_status,
                "Citation Contribution": citation_contribution,
                "Share of Citation Links": (
                    citation_contribution / total_citation_links
                    if total_citation_links
                    else 0.0
                ),
                "Distinct Citing Articles": distinct_citing_articles,
                "Average Citations per Citing Article": (
                    citation_contribution / distinct_citing_articles
                    if distinct_citing_articles
                    else 0.0
                ),
                "Distinct Source Papers Cited": distinct_source_papers,
                "Source Set Coverage": (
                    distinct_source_papers / total_source_papers
                    if total_source_papers
                    else 0.0
                ),
                "Grouping Basis": entry["grouping_basis"],
                "Matched Source ResearcherID(s)": matched_source_rids,
            }
        )

    columns = [
        "ResearcherID",
        "Preferred Name",
        "Source Set Author Match",
        "Citation Contribution",
        "Share of Citation Links",
        "Distinct Citing Articles",
        "Average Citations per Citing Article",
        "Distinct Source Papers Cited",
        "Source Set Coverage",
        "Grouping Basis",
        "Matched Source ResearcherID(s)",
    ]

    if not rows:
        return pd.DataFrame(columns=columns)

    return (
        pd.DataFrame(rows, columns=columns)
        .sort_values(
            [
                "Citation Contribution",
                "Distinct Citing Articles",
                "Preferred Name",
                "ResearcherID",
            ],
            ascending=[False, False, True, True],
        )
        .reset_index(drop=True)
    )

# =============================================================================
# Diagnostic tables
# =============================================================================

def build_citing_ut_summary(
    citation_links: List[Tuple[str, str]]
) -> pd.DataFrame:
    source_uids_by_citing_uid = defaultdict(set)

    for source_uid, citing_uid in citation_links:
        source_uids_by_citing_uid[citing_uid].add(source_uid)

    rows = []

    for citing_uid, source_uids in source_uids_by_citing_uid.items():
        rows.append(
            {
                "Citing UID": citing_uid,
                "Citation Link Count": len(source_uids),
                "Distinct Source Papers Cited": len(source_uids),
                "Source UIDs Cited": " ".join(sorted(source_uids)),
            }
        )

    if not rows:
        return pd.DataFrame(
            columns=[
                "Citing UID",
                "Citation Link Count",
                "Distinct Source Papers Cited",
                "Source UIDs Cited",
            ]
        )

    return (
        pd.DataFrame(rows)
        .sort_values(
            ["Citation Link Count", "Citing UID"],
            ascending=[False, True],
        )
        .reset_index(drop=True)
    )

# =============================================================================
# Excel output
# =============================================================================

def write_excel_workbook(
    excel_filename: str,
    source_author_df: pd.DataFrame,
    citing_author_df: pd.DataFrame,
    citing_ut_summary_df: pd.DataFrame,
    summary_df: pd.DataFrame,
) -> None:
    try:
        import xlsxwriter  # noqa: F401
        excel_engine = "xlsxwriter"
    except ImportError:
        excel_engine = "openpyxl"

    with pd.ExcelWriter(
        excel_filename,
        engine=excel_engine,
    ) as writer:

        source_author_df.to_excel(
            writer,
            index=False,
            sheet_name="Source Author Analysis",
        )

        citing_author_df.to_excel(
            writer,
            index=False,
            sheet_name="Citing Author Analysis",
        )

        citing_ut_summary_df.to_excel(
            writer,
            index=False,
            sheet_name="Citing UT Summary",
        )

        summary_df.to_excel(
            writer,
            index=False,
            header=False,
            sheet_name="Search Summary",
        )

        # Formatting is best with xlsxwriter; silently skip if openpyxl is used.
        try:
            workbook = writer.book

            header_fmt = workbook.add_format(
                {
                    "bold": True,
                    "border": 1,
                    "text_wrap": True,
                    "valign": "top",
                }
            )

            percent_fmt = workbook.add_format(
                {"num_format": "0.00%"}
            )

            decimal_fmt = workbook.add_format(
                {"num_format": "0.00"}
            )

            wrap_fmt = workbook.add_format(
                {"text_wrap": True, "valign": "top"}
            )

            sheet_tables = [
                ("Source Author Analysis", source_author_df),
                ("Citing Author Analysis", citing_author_df),
            ]

            sheet_tables.append(
                ("Citing UT Summary", citing_ut_summary_df)
            )

            for sheet_name, df_table in sheet_tables:
                ws = writer.sheets[sheet_name]
                ws.freeze_panes(1, 0)

                if len(df_table.columns):
                    ws.autofilter(
                        0,
                        0,
                        max(len(df_table), 1),
                        len(df_table.columns) - 1,
                    )

                for col_idx, col_name in enumerate(df_table.columns):
                    ws.write(0, col_idx, col_name, header_fmt)

                    sample_values = (
                        df_table[col_name]
                        .head(1000)
                        .astype(str)
                        .tolist()
                    )

                    max_len = max(
                        [len(str(col_name))]
                        + [len(value) for value in sample_values]
                    )

                    width = min(
                        max(12, max_len + 2),
                        60,
                    )

                    fmt = None

                    if col_name in {
                        "Percent of Source Papers",
                        "Share of Citation Links",
                        "Source Set Coverage",
                    }:
                        fmt = percent_fmt

                    if col_name == "Average Citations per Citing Article":
                        fmt = decimal_fmt
                        width = 18

                    if col_name == "Distinct Source Papers Cited":
                        width = 18

                    if col_name == "Source UIDs Cited":
                        width = 60
                        fmt = wrap_fmt

                    if col_name == "Matched Source ResearcherID(s)":
                        width = min(max(width, 28), 45)
                        fmt = wrap_fmt

                    ws.set_column(
                        col_idx,
                        col_idx,
                        width,
                        fmt,
                    )

            # Search Summary formatting
            ws_summary = writer.sheets["Search Summary"]
            ws_summary.set_column(0, 0, 38)
            ws_summary.set_column(1, 1, 100, wrap_fmt)

        except Exception:
            pass

# =============================================================================
# Main
# =============================================================================

def main() -> None:
    start_wall = datetime.datetime.now()
    load_dotenv()

    args = parse_args(sys.argv[1:])
    api_key = args.key or os.getenv("EXPANDED_APIKEY")

    if not api_key:
        print(
            "*** No API key supplied. Set EXPANDED_APIKEY in .env "
            "or use --key. ***"
        )
        sys.exit(1)

    run_params = dict(params)

    if args.query:
        run_params["usrQuery"] = args.query

    if not str(run_params.get("usrQuery") or "").strip():
        print(
            "*** No Web of Science query supplied. Set params['usrQuery'] "
            "in the script or use -q/--query. ***"
        )
        sys.exit(1)

    include_authors_without_rid = (
        INCLUDE_AUTHORS_WITHOUT_RID
        and not args.exclude_authors_without_rid
    )

    progress(f"Using query: {run_params['usrQuery']}")
    print(
        "Authors without ResearcherID: "
        + (
            "included with name fallback"
            if include_authors_without_rid
            else "excluded"
        )
    )

    # -------------------------------------------------------------------------
    # 1) Initial Short Record search
    # -------------------------------------------------------------------------

    stage_banner(1, "Retrieve source records using Short Record")

    try:
        source_records = wosesrclient_robust.get_all_records(
            api_key,
            run_params,
            run_params["firstRecord"],
            run_params["count"],
        )

    except (InvalidWoSQueryError, SRAuthError) as exc:
        print(exc)
        sys.exit(1)

    if not source_records:
        print("*** No records returned for this query. ***")
        sys.exit(0)

    # Preserve query order while deduping source records by UID.
    source_records_by_uid: Dict[str, Dict[str, Any]] = {}

    for rec in source_records:
        if not isinstance(rec, dict):
            continue

        uid = normalize_wos_uid(parse_uid(rec))

        if uid and uid not in source_records_by_uid:
            source_records_by_uid[uid] = rec

    source_uids = list(source_records_by_uid.keys())
    source_records = list(source_records_by_uid.values())

    if not source_uids:
        print("*** No source UIDs found. ***")
        sys.exit(0)

    progress(f"Retrieved {len(source_uids)} unique source records.")

    # -------------------------------------------------------------------------
    # 2) Analyze authors on the original source records
    # -------------------------------------------------------------------------

    stage_banner(2, "Analyze authors on source records")
    progress("Analyzing authors on source records...")

    source_author_df = build_source_author_analysis(
        source_records,
        include_authors_without_rid=include_authors_without_rid,
    )

    progress(
        f"Found {len(source_author_df)} grouped source-author identities."
    )

    # -------------------------------------------------------------------------
    # 3) Retrieve citing articles
    # -------------------------------------------------------------------------

    stage_banner(3, "Retrieve citing articles as Short Records")

    try:
        (
            citation_links,
            citing_records_by_uid,
            citing_errors,
        ) = retrieve_citation_links(
            api_key,
            source_uids,
        )

    except CitingAuthError as exc:
        print(exc)
        sys.exit(1)

    unique_citing_uids = list(
        dict.fromkeys(
            citing_uid
            for _, citing_uid in citation_links
        )
    )

    progress(
        f"Retrieved {len(citation_links)} citation links "
        f"from {len(unique_citing_uids)} unique citing articles."
    )

    # -------------------------------------------------------------------------
    # 4) Consolidate unique citing Short Records
    # -------------------------------------------------------------------------

    stage_banner(4, "Consolidate unique citing Short Records")

    missing_citing_sr_uids = [
        uid
        for uid in unique_citing_uids
        if uid not in citing_records_by_uid
    ]

    progress(
        f"Consolidated {len(citing_records_by_uid)} unique citing Short Records."
    )

    # -------------------------------------------------------------------------
    # 5) Analyze citing authors
    # -------------------------------------------------------------------------

    stage_banner(5, "Analyze citation contribution by citing author")
    progress("Analyzing citation contribution by citing author...")

    citing_author_df = build_citing_author_analysis(
        citing_records_by_uid,
        citation_links,
        source_records,
        include_authors_without_rid=include_authors_without_rid,
    )

    citing_ut_summary_df = build_citing_ut_summary(
        citation_links
    )

    progress(
        f"Found {len(citing_author_df)} grouped citing-author identities."
    )

    # -------------------------------------------------------------------------
    # 6) Write output
    # -------------------------------------------------------------------------

    stage_banner(6, "Write Excel output")

    raw_query = run_params["usrQuery"]
    safe_query = safe_slug_from_query(raw_query)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")

    if args.output:
        excel_filename = args.output

        if not excel_filename.lower().endswith(".xlsx"):
            excel_filename += ".xlsx"
    else:
        excel_filename = (
            f"WOS_CitingAuthorAnalysis_{safe_query}_{timestamp}.xlsx"
        )

    end_wall = datetime.datetime.now()
    elapsed_minutes = (
        end_wall - start_wall
    ).total_seconds() / 60

    links_without_sr = sum(
        1
        for _, citing_uid in citation_links
        if citing_uid not in citing_records_by_uid
    )

    source_match_counts = (
        citing_author_df["Source Set Author Match"].value_counts().to_dict()
        if "Source Set Author Match" in citing_author_df.columns
        else {}
    )

    summary_rows = [
        ["Original Query", raw_query],
        ["Run Started", f"{start_wall:%Y-%m-%d %H:%M:%S}"],
        ["Run Finished", f"{end_wall:%Y-%m-%d %H:%M:%S}"],
        ["Elapsed Minutes", f"{elapsed_minutes:.2f}"],
        ["", ""],
        ["Source Records Retrieved", len(source_uids)],
        ["Grouped Source Authors", len(source_author_df)],
        ["", ""],
        ["Citing Retrieval View", "Short Record (optionView=SR)"],
        ["Total Citation Links", len(citation_links)],
        ["Unique Citing Articles", len(unique_citing_uids)],
        [
            "Unique Citing Short Records Cached",
            len(citing_records_by_uid),
        ],
        [
            "Missing Citing Short Records",
            len(missing_citing_sr_uids),
        ],
        [
            "Citation Links Missing Author Metadata",
            links_without_sr,
        ],
        ["Grouped Citing Authors", len(citing_author_df)],
        [
            "Citing Authors - Exact Source Author",
            source_match_counts.get("Exact Source Author", 0),
        ],
        [
            "Citing Authors - Possible Source Author",
            source_match_counts.get("Possible Source Author", 0),
        ],
        [
            "Citing Authors - Not Matched to Source Set",
            source_match_counts.get("Not Matched to Source Set", 0),
        ],
        ["", ""],
        [
            "Citation Contribution Definition",
            (
                "Number of source-paper -> citing-paper citation links "
                "associated with an author. If one citing article cites five "
                "source papers, every author on that citing article receives "
                "a Citation Contribution of 5 from that article."
            ),
        ],
        [
            "Distinct Citing Articles",
            (
                "Number of unique citing papers on which the author appears. "
                "Included as diagnostic context; it is not the primary ranking."
            ),
        ],
        [
            "Distinct Source Papers Cited",
            (
                "Number of unique papers in the original source set reached "
                "by the author's citing articles."
            ),
        ],
        [
            "Average Citations per Citing Article",
            (
                "Citation Contribution divided by Distinct Citing Articles. "
                "This is the average number of source-set citations associated "
                "with each citing article on which the author appears."
            ),
        ],
        [
            "Source Set Coverage",
            (
                "Distinct Source Papers Cited divided by the total number of "
                "source records. This measures the breadth of the target set "
                "reached by the author's citing articles."
            ),
        ],
        [
            "Source Set Author Match",
            (
                "Exact Source Author means the same ResearcherID appears in the "
                "original source set. Possible Source Author means the ResearcherID "
                "differs or is missing but a punctuation/diacritic-normalized author "
                "name matches a source-set author. Possible matches are diagnostic "
                "and are not merged automatically."
            ),
        ],
        [
            "Author Grouping",
            (
                "Authors are grouped primarily by r_id and labeled with "
                "preferred_name.full_name. Authors lacking r_id are "
                + (
                    "grouped using a normalized-name fallback."
                    if include_authors_without_rid
                    else "excluded."
                )
            ),
        ],
        ["", ""],
        ["Citing Endpoint Retrieval Errors", len(citing_errors)],
    ]

    if citing_errors:
        summary_rows.append(["", ""])
        summary_rows.append(["Citing Endpoint Errors", ""])

        for message in citing_errors:
            summary_rows.append([message, ""])

    if missing_citing_sr_uids:
        summary_rows.append(["", ""])
        summary_rows.append(
            [
                "Missing Citing Short Record UTs",
                " ".join(missing_citing_sr_uids),
            ]
        )

    summary_df = pd.DataFrame(summary_rows)

    progress(f"Writing Excel workbook: {excel_filename}")

    write_excel_workbook(
        excel_filename=excel_filename,
        source_author_df=source_author_df,
        citing_author_df=citing_author_df,
        citing_ut_summary_df=citing_ut_summary_df,
        summary_df=summary_df,
    )

    print()
    progress("Analysis complete.")
    progress(f"Source records: {len(source_uids)}")
    progress(f"Citation links: {len(citation_links)}")
    progress(f"Unique citing articles: {len(unique_citing_uids)}")
    progress(f"Grouped source authors: {len(source_author_df)}")
    progress(f"Grouped citing authors: {len(citing_author_df)}")
    progress(f"Excel written to: {excel_filename}")

    progress(f"Elapsed: {elapsed_minutes:.2f} minutes")

if __name__ == "__main__":
    main()
