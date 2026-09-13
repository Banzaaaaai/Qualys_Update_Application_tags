"""Reconcile [VFZ] Applications & Platforms tags in Qualys (EU2 pod) with a CMDB Excel export.

This is architecturally derived from Qualys_Update_GAID_tags (update_gaid_tags.py) but is
NOT a GAID-tag script: application identity here is ASSET, not GAID. GAID is retained as
supporting metadata (description, reporting) and as a validated 1:1 companion key, never as
the primary matching key or tag name.

SOURCE OF TRUTH
  The CMDB Excel export is authoritative for which application tags exist under the
  "[VFZ] Applications & Platforms" parent, and for each application's desired IP scope.
  ASSET = application identity. GAID = supporting identifier, validated 1:1 with ASSET in
  preflight (fail closed on violation -- ambiguous identity must never reach Qualys).

WHAT IT DOES
  * Application in CMDB, no matching child tag             -> CREATE (NETWORK_RANGE), if it
    has usable IPs; otherwise reported MISSING_TAG_NO_IP_SCOPE and left uncreated.
  * Application in CMDB, tag exists as NETWORK_RANGE        -> ruleText replaced with EXACTLY
    the CMDB's desired IP set (complete replacement, never a merge).
  * Application in CMDB, tag exists as STATIC, has IPs      -> converted in place to
    NETWORK_RANGE, same tag id/name/parent.
  * Application in CMDB, tag exists as STATIC, no usable IPs -> left completely untouched.
  * Application in CMDB, tag exists as NAME_CONTAINS        -> NEVER written to, reported
    SKIPPED_UNSAFE_RULE_TYPE.
  * Tag under the parent with no matching CMDB application  -> DELETE candidate, gated behind
    --apply AND --allow-delete (default: report only).
  * A tag whose application still has resources, just none currently with usable IPs, is
    NEVER emptied or deleted -- see "NO USABLE IPS" below.

RESOURCE STATUS RULE (critical -- see README)
  Only rows with RESOURCE STATUS == "In Service" contribute IP addresses. ASSET STATUS
  (e.g. "Decommissioned") is informational context only and never by itself triggers
  deletion or scope changes: a decommissioned application can still have one live resource.

NO USABLE IPS
  An application with zero usable IPs is never treated as "delete this tag". A STATIC tag
  with no IPs is left alone; a NETWORK_RANGE tag is never cleared to an empty rule. Both are
  reported NO_USABLE_IPS. This is the single most important safety invariant here.

DELETE LOGIC
  A child tag is a delete candidate only when its application is genuinely absent from the
  CMDB, or (if ENABLE_DELETE_BY_ALL_RESOURCES_OOS) every single CMDB row for that ASSET --
  including rows with no IP -- has RESOURCE STATUS "Out of Service". A blank/unknown status
  on any row blocks deletion. Deletion additionally requires --apply AND --allow-delete.

TAG IDENTITY
  Existing tags are always updated in place (id, name, parent preserved); a normal update
  never deletes and recreates. Only writes the fields that actually change (ruleText-only,
  colour-only, or ruleType+ruleText together for a STATIC -> NETWORK_RANGE conversion).

Credentials come from QUALYS_USERNAME / QUALYS_PASSWORD env vars, falling back to
qualys_creds.txt next to this script (format: "key:\\tvalue" per line). Never accepted as a
CLI argument.

Default run mode is DRY-RUN. Pass --apply to write to Qualys, and additionally
--allow-delete to permit deletions. See README.md for the full rule set and rationale.
"""

from __future__ import annotations

import argparse
import csv
import datetime
import ipaddress
import json
import os
import re
import sys
import time
import unicodedata
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from xml.sax.saxutils import escape as xml_escape

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

# ==========================================================================
# Configuration
# ==========================================================================

QUALYS_BASE_URL = "https://qualysapi.qg2.apps.qualys.eu"
TAG_SEARCH_URL = f"{QUALYS_BASE_URL}/qps/rest/2.0/search/am/tag"
TAG_UPDATE_URL = f"{QUALYS_BASE_URL}/qps/rest/2.0/update/am/tag/{{tag_id}}"
TAG_CREATE_URL = f"{QUALYS_BASE_URL}/qps/rest/2.0/create/am/tag"
TAG_DELETE_URL = f"{QUALYS_BASE_URL}/qps/rest/2.0/delete/am/tag/{{tag_id}}"

TARGET_PARENT_TAG_NAME = "[VFZ] Applications & Platforms"
CHILD_TAG_PREFIX = "[VFZ] "

EXPECTED_RULE_TYPE = "NETWORK_RANGE"
STATIC_RULE_TYPES = {"", "STATIC"}
# Rule types this script recognizes but will never write to or delete,
# whatever the CMDB says -- they require a human to review and update
# manually. NAME_CONTAINS matches by hostname pattern; GROOVY is a scripted
# rule; ASSET_SEARCH matches by an arbitrary saved search (e.g. by QID).
# None of these have a safe automated translation to/from an IP-range rule.
NEVER_UPDATE_RULE_TYPES = {"NAME_CONTAINS", "GROOVY", "ASSET_SEARCH"}
KNOWN_RULE_TYPES = STATIC_RULE_TYPES | NEVER_UPDATE_RULE_TYPES | {EXPECTED_RULE_TYPE}

# Qualys does not publish an exact tag-name length ceiling for AssetView tags;
# this is a conservative assumed ceiling so a pathological CMDB value fails
# safely at plan time instead of producing a confusing API error at write time.
MAX_TAG_NAME_LENGTH = 256

# Qualys' documented CSAM limit on child tags per parent (see architecture doc).
MAX_CHILDREN_PER_PARENT = 350

EXCLUDED_IP_NETWORKS = ("169.254.0.0/16", "192.168.0.0/16")
_EXCLUDED_IP_RANGES = tuple(
    (int(net.network_address), int(net.broadcast_address))
    for net in (ipaddress.IPv4Network(c) for c in EXCLUDED_IP_NETWORKS)
)

# Safety cap on how many addresses a single "A-B" token may expand to, so a
# garbage/typo'd range can't blow up memory.
MAX_RANGE_SIZE = 1_000_000

# Only rows with this RESOURCE STATUS contribute IP addresses to an
# application's desired scope.
INCLUDE_RESOURCE_STATUSES = {"in service"}

# An application is only decommission-eligible-by-status when EVERY one of its
# CMDB rows (IP-bearing or not) carries a status in this set. A blank/unknown
# status is not in this set, so it blocks deletion by design.
DELETE_ON_RESOURCE_STATUSES = {"out of service"}
ENABLE_DELETE_BY_ALL_RESOURCES_OOS = True

DEFAULT_TAG_COLOR = "#0000FF"
ENABLE_COLOR_RECONCILIATION_DEFAULT = False
CREATE_TAGS_WITHOUT_IPS_DEFAULT = False

NEW_TAG_DESCRIPTION_TEMPLATE = "{asset} (GAID: {gaid})"
# Used to recover a probable GAID from a tag this script created earlier, to
# support optional --rename-tags matching when an ASSET's spelling changes.
DESCRIPTION_GAID_RE = re.compile(r"GAID:\s*([^)]+)\)")

PROTECTED_TAG_PATTERNS = [
    re.escape(TARGET_PARENT_TAG_NAME),
]

# Blast-radius guardrails. Exceeding any of these in --apply mode aborts
# before any mutation unless --force-large-change is also passed.
MAX_CREATE_CHANGES = 50
MAX_UPDATE_CHANGES = 250
MAX_DELETE_CHANGES = 20
MAX_TOTAL_CHANGES = 300
MAX_PERCENTAGE_CHANGED = 20  # percent of existing child tags

INITIAL_DELAY = 1
MAX_RETRIES = 5
REQUEST_TIMEOUT = 60
MAX_RATE_LIMIT_SLEEP = 300
TAG_SEARCH_PAGE_SIZE = 100

HTTP_HEADERS = {
    "Content-Type": "text/xml",
    "X-Requested-With": "update_application_tags.py",
}

IP_CELL_SPLIT_RE = re.compile(r"[,;\n\r]+")

# Header substrings (case-insensitive) used to infer the CMDB column mapping
# for the currently known schema. Adjust here if a real export's headers
# genuinely differ from these.
COLUMN_PATTERNS: dict[str, list[str]] = {
    "GAID": [r"^\s*gaid\s*$"],
    "ASSET": [r"^\s*asset\s*$"],
    "RVIT": [r"^\s*rvit\s*$"],
    "ASSET_STATUS": [r"asset\s*status"],
    "RESOURCE_NAME": [r"resource\s*name"],
    "FQDN": [r"^\s*fqdn\s*$"],
    "RESOURCE_TYPE": [r"resource\s*type"],
    "SYSTEM_SOFTWARE": [r"system\s*software"],
    "RESOURCE_STATUS": [r"resource\s*status"],
    "IPADDRESS": [r"ip\s*address(es)?", r"^\s*ip\s*$"],
    "ASSET_INSTANCE": [r"asset\s*instance"],
    "INSTANCE_ID": [r"instance\s*id"],
    "INSTANCE_STATUS": [r"instance\s*status"],
    "APPLICATION_COMPONENT": [r"application\s*component"],
    "COMPONENT_STATUS": [r"component\s*status"],
    "LOGICAL_CLUSTER": [r"logical\s*cluster"],
    "CLUSTER_STATUS": [r"cluster\s*status"],
    "TPL": [r"^\s*tpl\s*$"],
    "SQUAD": [r"^\s*squad\s*$"],
    "TRIBE": [r"^\s*tribe\s*$"],
}
REQUIRED_COLUMNS = ["ASSET", "GAID", "IPADDRESS", "RESOURCE_STATUS"]

REPORT_COLUMNS = [
    "action",
    "application",
    "gaid",
    "rvit",
    "asset_status",
    "resource_status_summary",
    "qualys_tag_id",
    "qualys_tag_name",
    "qualys_rule_type",
    "desired_rule_type",
    "old_ip_count",
    "new_ip_count",
    "ips_added_count",
    "ips_removed_count",
    "ips_excluded_count",
    "old_ip_summary",
    "new_ip_summary",
    "ips_added",
    "ips_removed",
    "reason",
    "verification_status",
    "error_message",
]
WRAP_COLUMNS = {
    "old_ip_summary",
    "new_ip_summary",
    "ips_added",
    "ips_removed",
    "reason",
    "error_message",
}
REPORT_COLUMN_WIDTHS = {
    "application": 32,
    "qualys_tag_name": 32,
    "old_ip_summary": 50,
    "new_ip_summary": 50,
    "ips_added": 34,
    "ips_removed": 34,
    "reason": 46,
    "error_message": 40,
}


# ==========================================================================
# Credentials
# ==========================================================================


class CredentialProvider:
    """Resolves Qualys credentials from environment first, then a local file.

    Passwords are never accepted as a CLI argument (they would be visible in
    shell history and process listings).
    """

    @staticmethod
    def get() -> tuple[str, str]:
        username = os.environ.get("QUALYS_USERNAME")
        password = os.environ.get("QUALYS_PASSWORD")
        if username and password:
            return username, password

        creds_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "qualys_creds.txt")
        if not os.path.isfile(creds_path):
            raise RuntimeError(
                "No credentials found: set QUALYS_USERNAME/QUALYS_PASSWORD or provide "
                "qualys_creds.txt"
            )

        creds: dict[str, str] = {}
        with open(creds_path, "r", encoding="utf-8") as f:
            for line in f:
                if ":" not in line:
                    continue
                key, _, value = line.partition(":")
                creds[key.strip().lower()] = value.strip()

        username = creds.get("user")
        password = creds.get("pass")
        if not username or not password:
            raise RuntimeError("qualys_creds.txt is missing 'user' and/or 'pass' entries")
        return username, password


# ==========================================================================
# Namespace-agnostic XML helpers
# ==========================================================================


def local_tag(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def find_child(elem, name):
    for child in elem:
        if local_tag(child.tag) == name:
            return child
    return None


def get_text(elem, name) -> str:
    child = find_child(elem, name)
    if child is None or child.text is None:
        return ""
    return child.text.strip()


# ==========================================================================
# IP / range normalization (ported from Qualys_Update_GAID_tags)
# ==========================================================================


def expand_entry_to_ints(token: str) -> set[int]:
    """Validate one token as an IPv4 address or 'A-B' range, return covered ints."""
    token = token.strip()
    if not token:
        raise ValueError("empty entry")

    if "-" in token:
        start_s, _, end_s = token.partition("-")
        start = ipaddress.IPv4Address(start_s.strip())
        end = ipaddress.IPv4Address(end_s.strip())
        if int(end) < int(start):
            raise ValueError(f"range end before start: {token}")
        if int(end) - int(start) + 1 > MAX_RANGE_SIZE:
            raise ValueError(f"range too large ({token})")
        return set(range(int(start), int(end) + 1))

    return {int(ipaddress.IPv4Address(token))}


def compact_ints_to_entries(int_set: set[int]) -> list[str]:
    """Merge address integers into the minimal sorted single-IP / 'A-B' entries."""
    if not int_set:
        return []
    ordered = sorted(int_set)
    runs = []
    start = prev = ordered[0]
    for n in ordered[1:]:
        if n == prev + 1:
            prev = n
            continue
        runs.append((start, prev))
        start = prev = n
    runs.append((start, prev))

    entries = []
    for a, b in runs:
        if a == b:
            entries.append(str(ipaddress.IPv4Address(a)))
        else:
            entries.append(f"{ipaddress.IPv4Address(a)}-{ipaddress.IPv4Address(b)}")
    return entries


def normalize_ip_set(raw_tokens) -> tuple[list[str], list[str]]:
    """Validate/union/recompact raw tokens. Returns (canonical_list, invalid_tokens)."""
    all_ints: set[int] = set()
    invalid = []
    for token in raw_tokens:
        if token is None:
            continue
        token = str(token).strip()
        if not token:
            continue
        try:
            all_ints |= expand_entry_to_ints(token)
        except ValueError:
            invalid.append(token)
            continue

    return compact_ints_to_entries(all_ints), invalid


def drop_excluded_networks(entries) -> tuple[list[str], list[str]]:
    """Strip EXCLUDED_IP_NETWORKS addresses from a canonical IP entry list.

    Applied to the CMDB's desired set only -- never to what Qualys currently
    holds -- so an excluded address already stored in a tag surfaces as a
    removal in the diff instead of being masked on both sides.
    """
    kept, dropped = set(), set()
    for entry in entries:
        for addr in expand_entry_to_ints(entry):
            if any(lo <= addr <= hi for lo, hi in _EXCLUDED_IP_RANGES):
                dropped.add(addr)
            else:
                kept.add(addr)
    return compact_ints_to_entries(kept), compact_ints_to_entries(dropped)


def parse_qualys_rule_text(rule_text: str) -> tuple[list[str], list[str]]:
    tokens = IP_CELL_SPLIT_RE.split(rule_text) if rule_text else []
    return normalize_ip_set(tokens)


def normalize_color(value) -> str:
    """Canonicalise a Qualys colour to '#RRGGBB'. Qualys renders it unpadded
    (e.g. '#0000FF' comes back as '#FF'), so both sides must go through this
    before comparing or every tag would appear to need a rewrite every run."""
    if not value:
        return ""
    text = str(value).strip().lstrip("#")
    if not text:
        return ""
    try:
        return f"#{int(text, 16):06X}"
    except ValueError:
        return ""


def normalize_asset_name(value) -> str:
    """Canonical application identity string: Unicode-normalized, invisible
    format characters stripped, tabs/newlines collapsed to spaces, internal
    whitespace runs collapsed, trimmed. Casing and all other visible
    characters are preserved -- the ASSET string is the authoritative
    tag-name source, not something to be silently lowercased.

    Real CMDB exports have been observed to carry zero-width spaces / BOM
    characters (copy/paste artifacts) inside ASSET values. Left in place
    these become invisible-but-real differences between otherwise-identical
    names (defeating matching) and can crash console output entirely on a
    Windows codepage that cannot render them. Unicode category "Cf" covers
    exactly this class of invisible formatting character.
    """
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Cf")
    text = text.replace("\t", " ").replace("\n", " ").replace("\r", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_gaid_key(value) -> str | None:
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if not text:
        return None
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def normalize_status(value) -> str:
    return str(value).strip().lower() if value is not None else ""


# ==========================================================================
# HTTP with retry (ported from Qualys_Update_GAID_tags)
# ==========================================================================


def request_with_retry(session, method, url, xml_body, label):
    delay = INITIAL_DELAY
    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.request(
                method,
                url,
                data=xml_body.encode("utf-8"),
                headers=HTTP_HEADERS,
                timeout=REQUEST_TIMEOUT,
            )
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
            last_error = str(exc)
            wait = min(delay, MAX_RATE_LIMIT_SLEEP)
            if attempt == MAX_RETRIES:
                break
            print(
                f"  [{label}] network error ({exc}); retrying in {wait}s "
                f"(attempt {attempt}/{MAX_RETRIES})"
            )
            time.sleep(wait)
            delay = min(delay * 2, MAX_RATE_LIMIT_SLEEP)
            continue

        if resp.status_code == 200:
            root = ET.fromstring(resp.text)
            response_code = get_text(root, "responseCode")
            if response_code and response_code != "SUCCESS":
                error_msg = get_text(root, "responseErrorDetails") or resp.text[:500]
                raise RuntimeError(f"{label}: Qualys API returned {response_code}: {error_msg}")
            return root

        if resp.status_code in (409, 429):
            wait_header = resp.headers.get("X-RateLimit-ToWait-Sec") or resp.headers.get(
                "Retry-After"
            )
            try:
                wait = (
                    min(float(wait_header), MAX_RATE_LIMIT_SLEEP)
                    if wait_header
                    else min(delay, MAX_RATE_LIMIT_SLEEP)
                )
            except ValueError:
                wait = min(delay, MAX_RATE_LIMIT_SLEEP)

            last_error = f"HTTP {resp.status_code} (rate-limited)"
            if attempt == MAX_RETRIES:
                break
            print(
                f"  [{label}] rate-limited (HTTP {resp.status_code}); waiting {wait}s "
                f"(attempt {attempt}/{MAX_RETRIES})"
            )
            time.sleep(wait)
            delay = min(delay * 2, MAX_RATE_LIMIT_SLEEP)
            continue

        last_error = f"HTTP {resp.status_code}: {resp.text[:500]}"
        wait = min(delay, MAX_RATE_LIMIT_SLEEP)
        if attempt == MAX_RETRIES:
            break
        print(f"  [{label}] {last_error}; retrying in {wait}s (attempt {attempt}/{MAX_RETRIES})")
        time.sleep(wait)
        delay = min(delay * 2, MAX_RATE_LIMIT_SLEEP)

    raise RuntimeError(f"{label}: giving up after {MAX_RETRIES} attempts: {last_error}")


# ==========================================================================
# QualysTagClient
# ==========================================================================


@dataclass
class QualysTag:
    tag_id: str
    tag_name: str
    parent_tag_id: str
    rule_type: str
    rule_text: str
    color: str
    description: str
    criticality: str
    created: str
    modified: str

    @property
    def gaid_from_description(self) -> str | None:
        match = DESCRIPTION_GAID_RE.search(self.description or "")
        return match.group(1).strip() if match else None


def _parse_tag_elements(data_elem) -> list[QualysTag]:
    tags = []
    if data_elem is None:
        return tags
    for tag_elem in data_elem:
        if local_tag(tag_elem.tag) != "Tag":
            continue
        tags.append(
            QualysTag(
                tag_id=get_text(tag_elem, "id"),
                tag_name=get_text(tag_elem, "name"),
                parent_tag_id=get_text(tag_elem, "parentTagId"),
                rule_type=get_text(tag_elem, "ruleType"),
                rule_text=get_text(tag_elem, "ruleText"),
                color=get_text(tag_elem, "color"),
                description=get_text(tag_elem, "description"),
                criticality=get_text(tag_elem, "criticalityScore"),
                created=get_text(tag_elem, "created"),
                modified=get_text(tag_elem, "modified"),
            )
        )
    return tags


class QualysTagClient:
    """Thin wrapper over the am/tag REST endpoints: search/create/update/delete,
    with deterministic id-cursor pagination and rate-limit-aware retry."""

    def __init__(self, session):
        self.session = session

    def _search(self, xml_body, label) -> list[QualysTag]:
        root = request_with_retry(self.session, "POST", TAG_SEARCH_URL, xml_body, label)
        return _parse_tag_elements(find_child(root, "data"))

    def _search_paginated(self, criteria_xml_fn, label) -> list[QualysTag]:
        all_tags: list[QualysTag] = []
        last_id = 0
        page_num = 0
        while True:
            page_num += 1
            xml_body = (
                "<ServiceRequest>"
                f"<preferences><limitResults>{TAG_SEARCH_PAGE_SIZE}</limitResults></preferences>"
                f"<filters>{criteria_xml_fn(last_id)}</filters>"
                "</ServiceRequest>"
            )
            root = request_with_retry(
                self.session, "POST", TAG_SEARCH_URL, xml_body, f"{label} page {page_num}"
            )
            page_tags = _parse_tag_elements(find_child(root, "data"))
            all_tags.extend(page_tags)

            has_more = get_text(root, "hasMoreRecords").lower() == "true"
            last_id_text = get_text(root, "lastId")
            if last_id_text:
                last_id = int(last_id_text)
            else:
                numeric_ids = [int(t.tag_id) for t in page_tags if t.tag_id.isdigit()]
                if numeric_ids:
                    last_id = max(last_id, max(numeric_ids))

            print(f"  [{label} page {page_num}] fetched {len(page_tags)} tag(s) (last_id={last_id})")
            if not has_more or not page_tags:
                break
            time.sleep(INITIAL_DELAY)
        return all_tags

    def find_by_name(self, name: str) -> list[QualysTag]:
        xml_body = (
            "<ServiceRequest><filters>"
            f'<Criteria field="name" operator="EQUALS">{xml_escape(name)}</Criteria>'
            "</filters></ServiceRequest>"
        )
        return self._search(xml_body, f"find tag by name {name!r}")

    def fetch_by_id(self, tag_id) -> QualysTag | None:
        xml_body = (
            "<ServiceRequest><filters>"
            f'<Criteria field="id" operator="EQUALS">{xml_escape(str(tag_id))}</Criteria>'
            "</filters></ServiceRequest>"
        )
        tags = self._search(xml_body, f"verify tag {tag_id}")
        return tags[0] if tags else None

    def fetch_by_name(self, name: str) -> QualysTag | None:
        tags = self.find_by_name(name)
        return tags[0] if tags else None

    def fetch_all(self) -> list[QualysTag]:
        """Enumerate every tag in the tenant via plain id-cursor pagination.

        Qualys' am/tag search does not accept "parentTagId" as a filter
        field (confirmed live: INVALID_REQUEST names the valid set as
        parent, color, provider, ruleType, name, id, criticalityScore) and
        the semantics of "parent" there are not documented clearly enough to
        rely on. Fetching everything and filtering by parent_tag_id
        client-side, as the GAID script does, is the proven-safe approach.
        """

        def criteria(last_id):
            return f'<Criteria field="id" operator="GREATER">{last_id}</Criteria>'

        return self._search_paginated(criteria, "search all tags")

    def fetch_children(self, parent_tag_id) -> list[QualysTag]:
        return [t for t in self.fetch_all() if t.parent_tag_id == str(parent_tag_id)]

    def update(self, tag_id, fields_xml, label) -> None:
        xml_body = f"<ServiceRequest><data><Tag>{fields_xml}</Tag></data></ServiceRequest>"
        request_with_retry(self.session, "POST", TAG_UPDATE_URL.format(tag_id=tag_id), xml_body, label)

    def create(self, fields_xml, label) -> None:
        xml_body = f"<ServiceRequest><data><Tag>{fields_xml}</Tag></data></ServiceRequest>"
        request_with_retry(self.session, "POST", TAG_CREATE_URL, xml_body, label)

    def delete(self, tag_id, label) -> None:
        request_with_retry(
            self.session, "POST", TAG_DELETE_URL.format(tag_id=tag_id), "<ServiceRequest></ServiceRequest>", label
        )


# ==========================================================================
# QualysTagModel -- current state of the target hierarchy
# ==========================================================================


def normalize_for_case_match(name: str) -> str:
    """Casefold + whitespace-collapse a tag/asset name for identity matching
    that is insensitive to case and incidental whitespace differences (e.g.
    'Cyberark' vs 'CyberArk', 'I&M portal' vs 'I&M Portal'). This is used
    only to FIND the right existing tag; the tag's own name is left exactly
    as-is unless --rename-tags is also passed -- ASSET casing is CMDB
    metadata, not grounds for an automatic rename."""
    return re.sub(r"\s+", " ", name.strip()).casefold()


class QualysTagModel:
    """Holds the resolved parent tag and its direct child tags."""

    def __init__(self, parent: QualysTag, children: list[QualysTag]):
        self.parent = parent
        self.children = children
        self.by_name: dict[str, QualysTag] = {t.tag_name: t for t in children}
        self.by_name_casefold: dict[str, list[QualysTag]] = {}
        for t in children:
            self.by_name_casefold.setdefault(normalize_for_case_match(t.tag_name), []).append(t)
        self.by_gaid: dict[str, QualysTag] = {}
        for t in children:
            gaid = t.gaid_from_description
            if gaid:
                self.by_gaid.setdefault(gaid, t)

    @classmethod
    def resolve(cls, client: QualysTagClient) -> "QualysTagModel":
        candidates = client.find_by_name(TARGET_PARENT_TAG_NAME)
        if len(candidates) == 0:
            raise RuntimeError(
                f"Parent tag {TARGET_PARENT_TAG_NAME!r} not found in tenant -- aborting "
                "before any mutation."
            )
        if len(candidates) > 1:
            ids = ", ".join(t.tag_id for t in candidates)
            raise RuntimeError(
                f"Multiple tags named {TARGET_PARENT_TAG_NAME!r} found (ids: {ids}) -- "
                "identity is ambiguous, aborting before any mutation."
            )
        parent = candidates[0]
        children = client.fetch_children(parent.tag_id)
        return cls(parent, children)


# ==========================================================================
# CMDB reading
# ==========================================================================


def find_column(headers, patterns) -> int | None:
    for pattern in patterns:
        regex = re.compile(pattern, re.IGNORECASE)
        for idx, header in enumerate(headers):
            if header and regex.search(str(header)):
                return idx
    return None


class CmdbReader:
    """Loads the CMDB Excel export and resolves the known column layout."""

    def __init__(self, path: str):
        self.path = path

    def load(self):
        if not os.path.isfile(self.path):
            raise RuntimeError(f"Input file not found: {self.path}")
        wb = load_workbook(self.path, data_only=True, read_only=True)
        if not wb.sheetnames:
            raise RuntimeError(f"Workbook {self.path!r} has no worksheets")
        sheet_name = wb.sheetnames[0]
        ws = wb[sheet_name]
        row_iter = ws.iter_rows(values_only=True)
        try:
            headers = list(next(row_iter))
        except StopIteration:
            raise RuntimeError(f"Sheet {sheet_name!r} is empty")

        rows = []
        for row in row_iter:
            if row is None or all(v is None for v in row):
                continue
            rows.append(row)

        print(f"=== CMDB discovery: {self.path} ===")
        print(f"Sheets found: {wb.sheetnames} (using {sheet_name!r})")
        print(f"Header row: {headers}")
        for row in rows[:3]:
            print(f"  sample: {row}")

        columns = self._resolve_columns(headers)
        print()
        print("=== Resolved column mapping ===")
        for logical, idx in sorted(columns.items()):
            print(f"  {logical:24s} -> {headers[idx]!r} (index {idx})")
        missing_optional = [c for c in COLUMN_PATTERNS if c not in columns]
        if missing_optional:
            print(f"  (not found / optional: {', '.join(sorted(missing_optional))})")
        print()

        return headers, rows, columns

    @staticmethod
    def _resolve_columns(headers) -> dict[str, int]:
        columns: dict[str, int] = {}
        for logical, patterns in COLUMN_PATTERNS.items():
            idx = find_column(headers, patterns)
            if idx is not None:
                columns[logical] = idx

        missing_required = [c for c in REQUIRED_COLUMNS if c not in columns]
        if missing_required:
            raise RuntimeError(
                "Could not resolve required column(s) "
                f"{missing_required} from header {headers!r}. COLUMN_PATTERNS at the top "
                "of this script may need adjusting for this export's real headers."
            )
        return columns

    @staticmethod
    def cell(row, columns, logical, default=None):
        idx = columns.get(logical)
        if idx is None or idx >= len(row):
            return default
        value = row[idx]
        return default if value is None else value


# ==========================================================================
# Application model
# ==========================================================================


@dataclass
class ApplicationRecord:
    asset: str  # canonical display name (also the tag-name suffix)
    gaid: str = ""
    rvit: str = ""
    asset_status_labels: set[str] = field(default_factory=set)
    all_row_statuses: list[str] = field(default_factory=list)  # every row, for delete-gating
    ip_bearing_row_statuses: list[str] = field(default_factory=list)  # rows with a non-blank IP cell
    in_service_ip_tokens: list[str] = field(default_factory=list)

    # Computed by DesiredStateBuilder.finalize()
    desired_ips: list[str] = field(default_factory=list)
    excluded_ips: list[str] = field(default_factory=list)
    invalid_ip_tokens: list[str] = field(default_factory=list)
    had_ips_before_exclusion: bool = False

    @property
    def resource_status_summary(self) -> str:
        counts: dict[str, int] = {}
        for label in self.ip_bearing_row_statuses:
            counts[label] = counts.get(label, 0) + 1
        ordered = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))
        return ", ".join(f"{label}: {n}" for label, n in ordered)

    @property
    def all_resources_out_of_service(self) -> bool:
        labels = {normalize_status(s) for s in self.all_row_statuses}
        return bool(labels) and labels <= DELETE_ON_RESOURCE_STATUSES

    @property
    def asset_status_display(self) -> str:
        return ", ".join(sorted(s for s in self.asset_status_labels if s))


@dataclass
class ValidationError:
    row_index: int
    message: str


class DesiredStateBuilder:
    """Builds the per-ASSET ApplicationRecord map from CMDB rows, and validates
    the GAID<->ASSET 1:1 invariant this script's identity model depends on."""

    def __init__(self, headers, rows, columns):
        self.headers = headers
        self.rows = rows
        self.columns = columns
        self.errors: list[ValidationError] = []

    def build(self) -> dict[str, ApplicationRecord]:
        records: dict[str, ApplicationRecord] = {}
        asset_to_gaids: dict[str, set[str]] = {}
        gaid_to_assets: dict[str, set[str]] = {}

        for idx, row in enumerate(self.rows, start=2):  # +2: header is row 1, 1-indexed
            raw_asset = CmdbReader.cell(row, self.columns, "ASSET")
            raw_gaid = CmdbReader.cell(row, self.columns, "GAID")
            asset = normalize_asset_name(raw_asset)
            gaid = normalize_gaid_key(raw_gaid) or ""

            if not asset and not gaid:
                continue  # a genuinely blank data row; not an application row
            if not asset:
                self.errors.append(
                    ValidationError(idx, f"blank ASSET on row with GAID={raw_gaid!r}")
                )
                continue

            if asset not in records:
                records[asset] = ApplicationRecord(asset=asset)
            rec = records[asset]

            if gaid:
                if not rec.gaid:
                    rec.gaid = gaid
                elif rec.gaid != gaid:
                    self.errors.append(
                        ValidationError(
                            idx,
                            f"ASSET {asset!r} maps to multiple GAIDs: {rec.gaid!r} and {gaid!r}",
                        )
                    )
                asset_to_gaids.setdefault(asset, set()).add(gaid)
                gaid_to_assets.setdefault(gaid, set()).add(asset)

            rvit = CmdbReader.cell(row, self.columns, "RVIT")
            if rvit is not None and not rec.rvit:
                rec.rvit = str(rvit).strip()

            asset_status = CmdbReader.cell(row, self.columns, "ASSET_STATUS")
            if asset_status is not None and str(asset_status).strip():
                rec.asset_status_labels.add(str(asset_status).strip())

            resource_status_raw = CmdbReader.cell(row, self.columns, "RESOURCE_STATUS")
            resource_status_label = (
                str(resource_status_raw).strip() if resource_status_raw is not None else "(blank)"
            )
            rec.all_row_statuses.append(resource_status_label)

            ip_cell = CmdbReader.cell(row, self.columns, "IPADDRESS")
            if ip_cell is not None:
                rec.ip_bearing_row_statuses.append(resource_status_label)
                if normalize_status(resource_status_raw) in INCLUDE_RESOURCE_STATUSES:
                    rec.in_service_ip_tokens.extend(IP_CELL_SPLIT_RE.split(str(ip_cell)))

        # 1:1 GAID<->ASSET validation -- fail closed on any violation.
        for asset, gaids in asset_to_gaids.items():
            if len(gaids) > 1:
                self.errors.append(
                    ValidationError(0, f"ASSET {asset!r} maps to multiple GAIDs: {sorted(gaids)}")
                )
        for gaid, assets in gaid_to_assets.items():
            if len(assets) > 1:
                self.errors.append(
                    ValidationError(0, f"GAID {gaid!r} maps to multiple ASSETs: {sorted(assets)}")
                )

        for rec in records.values():
            new_ips, invalid = normalize_ip_set(rec.in_service_ip_tokens)
            rec.invalid_ip_tokens = invalid
            rec.had_ips_before_exclusion = bool(new_ips)
            kept, excluded = drop_excluded_networks(new_ips)
            rec.desired_ips = kept
            rec.excluded_ips = excluded

        return records


# ==========================================================================
# Change planning (diff engine)
# ==========================================================================

ACTION_NO_CHANGE = "NO_CHANGE"
ACTION_UPDATE = "UPDATE"
ACTION_CONVERT = "CONVERT_STATIC_TO_DYNAMIC"
ACTION_CREATE = "CREATE"
ACTION_DELETE = "DELETE"
ACTION_SKIP_NO_IPS = "SKIP_NO_USABLE_IPS"
ACTION_SKIP_UNSUPPORTED = "SKIP_UNSUPPORTED_RULE_TYPE"
ACTION_SKIP_AMBIGUOUS = "SKIP_AMBIGUOUS_SOURCE"
ACTION_ERROR = "ERROR"


def _new_row(**kwargs) -> dict:
    row = {col: "" for col in REPORT_COLUMNS}
    row.update(kwargs)
    return row


class ChangePlanner:
    """Produces one report row (+ internal execution fields, keys prefixed
    with '_') per application, plus one per orphaned Qualys child tag."""

    def __init__(
        self,
        applications: dict[str, ApplicationRecord],
        tag_model: QualysTagModel,
        manage_color: bool,
        create_tags_without_ips: bool,
        rename_tags: bool,
    ):
        self.applications = applications
        self.tag_model = tag_model
        self.manage_color = manage_color
        self.create_tags_without_ips = create_tags_without_ips
        self.rename_tags = rename_tags

    def plan(self) -> list[dict]:
        rows: list[dict] = []
        claimed_tag_ids: set[str] = set()

        for asset in sorted(self.applications):
            app = self.applications[asset]
            row = self._plan_one_application(app, claimed_tag_ids)
            rows.append(row)

        for tag in self.tag_model.children:
            if tag.tag_id in claimed_tag_ids:
                continue
            rows.append(self._plan_orphan_tag(tag))

        return rows

    # -- per application -----------------------------------------------

    def _plan_one_application(self, app: ApplicationRecord, claimed_tag_ids: set[str]) -> dict:
        desired_name = CHILD_TAG_PREFIX + app.asset
        common = dict(
            application=app.asset,
            gaid=app.gaid,
            rvit=app.rvit,
            asset_status=app.asset_status_display,
            resource_status_summary=app.resource_status_summary,
            desired_rule_type=EXPECTED_RULE_TYPE,
            ips_excluded_count=len(app.excluded_ips),
        )

        if app.invalid_ip_tokens:
            return _new_row(
                action=ACTION_ERROR,
                verification_status="NOT_APPLICABLE",
                reason="invalid IP/range entr" + ("y" if len(app.invalid_ip_tokens) == 1 else "ies"),
                error_message="invalid tokens: " + ", ".join(app.invalid_ip_tokens),
                **common,
            )

        if len(desired_name) > MAX_TAG_NAME_LENGTH:
            return _new_row(
                action=ACTION_ERROR,
                verification_status="NOT_APPLICABLE",
                reason="ASSET name produces a tag name exceeding the assumed Qualys length limit",
                error_message=f"desired tag name is {len(desired_name)} chars (limit {MAX_TAG_NAME_LENGTH})",
                **common,
            )

        tag = self.tag_model.by_name.get(desired_name)
        match_note = ""

        # Case/whitespace-insensitive fallback: CMDB spelling drifts (e.g.
        # "Cyberark" -> "CyberArk") must not read as "old app gone, new app
        # appeared". The existing tag is matched and its IP scope
        # reconciled; its name is left exactly as-is unless --rename-tags is
        # also passed (see the "renamed" handling in _plan_existing).
        if tag is None:
            norm_key = normalize_for_case_match(desired_name)
            unclaimed = [
                t for t in self.tag_model.by_name_casefold.get(norm_key, [])
                if t.tag_id not in claimed_tag_ids
            ]
            if len(unclaimed) == 1:
                tag = unclaimed[0]
                match_note = (
                    f"matched case-insensitively (existing tag name {tag.tag_name!r} differs "
                    f"from CMDB ASSET only by case/whitespace)"
                )
            elif len(unclaimed) > 1:
                return _new_row(
                    action=ACTION_ERROR,
                    verification_status="NOT_APPLICABLE",
                    reason="multiple existing tags match this application case-insensitively; ambiguous",
                    error_message="candidates: " + ", ".join(t.tag_name for t in unclaimed),
                    **common,
                )

        # Optional rename-safe matching via a GAID recorded in an existing
        # tag's description, when neither exact nor case-insensitive
        # matching found anything.
        if tag is None and self.rename_tags and app.gaid:
            candidate = self.tag_model.by_gaid.get(app.gaid)
            if candidate is not None and candidate.tag_id not in claimed_tag_ids:
                tag = candidate

        if tag is not None:
            claimed_tag_ids.add(tag.tag_id)

        if tag is None:
            return self._plan_create(app, desired_name, common)
        return self._plan_existing(app, tag, desired_name, common, match_note)

    def _plan_create(self, app: ApplicationRecord, desired_name: str, common: dict) -> dict:
        if app.desired_ips:
            description = NEW_TAG_DESCRIPTION_TEMPLATE.format(asset=app.asset, gaid=app.gaid or "unknown")
            return _new_row(
                action=ACTION_CREATE,
                qualys_tag_name=desired_name,
                qualys_rule_type="(none)",
                new_ip_count=len(app.desired_ips),
                new_ip_summary=", ".join(app.desired_ips),
                ips_added=", ".join(app.desired_ips),
                ips_added_count=len(app.desired_ips),
                reason="application present in CMDB with usable IPs, no matching tag",
                verification_status="PENDING",
                _new_ips=app.desired_ips,
                _description=description,
                _new_color=DEFAULT_TAG_COLOR,
                **common,
            )
        if self.create_tags_without_ips:
            return _new_row(
                action=ACTION_CREATE,
                qualys_tag_name=desired_name,
                qualys_rule_type="(none)",
                reason="--create-static-for-no-ip: creating placeholder STATIC tag",
                verification_status="PENDING",
                _new_ips=[],
                _description=NEW_TAG_DESCRIPTION_TEMPLATE.format(asset=app.asset, gaid=app.gaid or "unknown"),
                _create_static=True,
                _new_color=DEFAULT_TAG_COLOR,
                **common,
            )
        return _new_row(
            action=ACTION_SKIP_NO_IPS,
            qualys_tag_name=desired_name,
            reason="MISSING_TAG_NO_IP_SCOPE: no tag exists and CMDB has no usable IPs",
            verification_status="NOT_APPLICABLE",
            **common,
        )

    def _plan_existing(
        self, app: ApplicationRecord, tag: QualysTag, desired_name: str, common: dict, match_note: str = ""
    ) -> dict:
        old_ips, _ = parse_qualys_rule_text(tag.rule_text)
        base = dict(
            qualys_tag_id=tag.tag_id,
            qualys_tag_name=tag.tag_name,
            qualys_rule_type=tag.rule_type or "(static)",
            old_ip_summary=", ".join(old_ips),
            old_ip_count=len(old_ips),
            **common,
        )
        color_differs = self.manage_color and normalize_color(tag.color) != normalize_color(DEFAULT_TAG_COLOR)
        renamed = tag.tag_name != desired_name

        def with_note(reason: str) -> str:
            return f"{reason}; {match_note}" if match_note else reason

        if tag.rule_type in NEVER_UPDATE_RULE_TYPES:
            return _new_row(
                action=ACTION_SKIP_UNSUPPORTED,
                reason=with_note(
                    f"SKIPPED_UNSAFE_RULE_TYPE: {tag.rule_type} tags are never written to "
                    "automatically -- requires manual review"
                ),
                verification_status="NOT_APPLICABLE",
                **base,
            )

        if tag.rule_type not in KNOWN_RULE_TYPES:
            return _new_row(
                action=ACTION_ERROR,
                reason=with_note("unexpected ruleType; failing safe at tag level"),
                error_message=f"ruleType {tag.rule_type!r} is not one of {sorted(KNOWN_RULE_TYPES)}",
                verification_status="NOT_APPLICABLE",
                **base,
            )

        is_static = tag.rule_type in STATIC_RULE_TYPES

        if is_static and not app.desired_ips:
            reason = "NO_USABLE_IPS: static tag preserved unchanged"
            if renamed and self.rename_tags:
                reason += f"; would also rename to {desired_name!r}"
            return _new_row(
                action=ACTION_UPDATE if (renamed and self.rename_tags) or color_differs else ACTION_NO_CHANGE,
                reason=with_note(reason),
                verification_status="NOT_APPLICABLE" if not ((renamed and self.rename_tags) or color_differs) else "PENDING",
                _rename_only=renamed and self.rename_tags,
                _new_name=desired_name if (renamed and self.rename_tags) else "",
                _color_only=not (renamed and self.rename_tags),
                _new_color=DEFAULT_TAG_COLOR if color_differs else "",
                _tag=tag,
                **base,
            )

        if is_static and app.desired_ips:
            return _new_row(
                action=ACTION_CONVERT,
                new_ip_summary=", ".join(app.desired_ips),
                new_ip_count=len(app.desired_ips),
                ips_added=", ".join(app.desired_ips),
                ips_added_count=len(app.desired_ips),
                reason=with_note(
                    "static tag converted in place to NETWORK_RANGE (CMDB has usable IPs)"
                    + (f"; renamed to {desired_name!r}" if renamed and self.rename_tags else "")
                ),
                verification_status="PENDING",
                _new_ips=app.desired_ips,
                _is_conversion=True,
                _new_name=desired_name if (renamed and self.rename_tags) else "",
                _new_color=DEFAULT_TAG_COLOR,
                _tag=tag,
                **base,
            )

        # NETWORK_RANGE tag.
        if not app.desired_ips:
            reason = "NO_USABLE_IPS: existing NETWORK_RANGE rule left as-is, not cleared"
            do_something = color_differs or (renamed and self.rename_tags)
            return _new_row(
                action=ACTION_UPDATE if do_something else ACTION_NO_CHANGE,
                reason=with_note(reason),
                verification_status="PENDING" if do_something else "NOT_APPLICABLE",
                _rename_only=renamed and self.rename_tags and not color_differs,
                _color_only=color_differs,
                _new_name=desired_name if (renamed and self.rename_tags) else "",
                _new_color=DEFAULT_TAG_COLOR if color_differs else "",
                _tag=tag,
                **base,
            )

        old_set, new_set = set(old_ips), set(app.desired_ips)
        added = sorted(new_set - old_set, key=lambda c: app.desired_ips.index(c) if c in app.desired_ips else 0)
        removed = sorted(old_set - new_set, key=lambda c: old_ips.index(c) if c in old_ips else 0)
        has_diff = old_set != new_set or (renamed and self.rename_tags) or color_differs

        return _new_row(
            action=ACTION_UPDATE if has_diff else ACTION_NO_CHANGE,
            new_ip_summary=", ".join(app.desired_ips),
            new_ip_count=len(app.desired_ips),
            ips_added=", ".join(added),
            ips_added_count=len(added),
            ips_removed=", ".join(removed),
            ips_removed_count=len(removed),
            reason=with_note(
                "IP set differs from CMDB" if old_set != new_set else (
                    "colour/name reconciliation only" if has_diff else "already matches CMDB"
                )
            ),
            verification_status="PENDING" if has_diff else "NOT_APPLICABLE",
            _new_ips=app.desired_ips,
            _new_name=desired_name if (renamed and self.rename_tags) else "",
            _new_color=DEFAULT_TAG_COLOR if color_differs else "",
            _tag=tag,
            **base,
        )

    # -- tags present in Qualys but not matched to any application ------

    def _plan_orphan_tag(self, tag: QualysTag) -> dict:
        old_ips, _ = parse_qualys_rule_text(tag.rule_text)
        base = dict(
            qualys_tag_id=tag.tag_id,
            qualys_tag_name=tag.tag_name,
            qualys_rule_type=tag.rule_type or "(static)",
            old_ip_summary=", ".join(old_ips),
            old_ip_count=len(old_ips),
        )

        for pattern in PROTECTED_TAG_PATTERNS:
            if re.fullmatch(pattern, tag.tag_name):
                return _new_row(
                    action=ACTION_NO_CHANGE,
                    reason="protected tag name pattern match; never a deletion candidate",
                    verification_status="NOT_APPLICABLE",
                    **base,
                )

        # Rule-type safety applies to deletion exactly as it applies to
        # writes: NAME_CONTAINS / GROOVY / ASSET_SEARCH tags (and any
        # unrecognized type) are never removed just because their name
        # doesn't literally appear in this month's CMDB -- the CMDB may
        # represent the same application differently, and these rule types
        # are never guessed at; a human must review and update them.
        if tag.rule_type in NEVER_UPDATE_RULE_TYPES:
            return _new_row(
                action=ACTION_SKIP_UNSUPPORTED,
                reason=f"SKIPPED_UNSAFE_RULE_TYPE: {tag.rule_type} tags are never deleted "
                "automatically, even when their name has no literal CMDB match -- requires "
                "manual review",
                verification_status="NOT_APPLICABLE",
                **base,
            )
        if tag.rule_type not in KNOWN_RULE_TYPES:
            return _new_row(
                action=ACTION_ERROR,
                reason="unexpected ruleType; failing safe rather than guessing whether deletion is safe",
                error_message=f"ruleType {tag.rule_type!r} is not one of {sorted(KNOWN_RULE_TYPES)}",
                verification_status="NOT_APPLICABLE",
                **base,
            )

        if not tag.tag_name.startswith(CHILD_TAG_PREFIX):
            return _new_row(
                action=ACTION_SKIP_AMBIGUOUS,
                reason="tag name does not follow the expected '[VFZ] <ASSET>' pattern; left alone",
                verification_status="NOT_APPLICABLE",
                **base,
            )

        asset_guess = tag.tag_name[len(CHILD_TAG_PREFIX):]
        app = self.applications.get(asset_guess)

        if app is None:
            return _new_row(
                action=ACTION_DELETE,
                application=asset_guess,
                ips_removed=", ".join(old_ips),
                ips_removed_count=len(old_ips),
                reason="application absent from CMDB",
                verification_status="PENDING",
                _tag=tag,
                **base,
            )

        if ENABLE_DELETE_BY_ALL_RESOURCES_OOS and app.all_resources_out_of_service:
            return _new_row(
                action=ACTION_DELETE,
                application=asset_guess,
                gaid=app.gaid,
                asset_status=app.asset_status_display,
                resource_status_summary=app.resource_status_summary,
                ips_removed=", ".join(old_ips),
                ips_removed_count=len(old_ips),
                reason="all CMDB resources for this application are Out of Service",
                verification_status="PENDING",
                _tag=tag,
                **base,
            )

        # The application exists in the CMDB but this tag was not claimed by
        # the forward pass -- most likely a data-quality mismatch (e.g. the
        # matched-by-name tag differs textually from the recovered guess).
        return _new_row(
            action=ACTION_NO_CHANGE,
            application=asset_guess,
            gaid=app.gaid,
            asset_status=app.asset_status_display,
            resource_status_summary=app.resource_status_summary,
            reason="present in CMDB; not deletion-eligible; left unmodified",
            verification_status="NOT_APPLICABLE",
            **base,
        )


# ==========================================================================
# Preflight
# ==========================================================================


class PreflightError(RuntimeError):
    pass


class PreflightValidator:
    @staticmethod
    def validate_source(builder: DesiredStateBuilder) -> None:
        if builder.errors:
            print("=== PREFLIGHT FAILED: source validation errors ===")
            for err in builder.errors[:50]:
                where = f"row {err.row_index}" if err.row_index else "(aggregate)"
                print(f"  [{where}] {err.message}")
            if len(builder.errors) > 50:
                print(f"  ... and {len(builder.errors) - 50} more")
            raise PreflightError(
                f"{len(builder.errors)} source validation error(s); see above. "
                "No Qualys mutation will be attempted."
            )

    @staticmethod
    def validate_plan(
        plan_rows: list[dict],
        existing_child_count: int,
        apply_mode: bool,
        allow_delete: bool,
        max_create: int,
        max_update: int,
        max_delete: int,
        max_total: int,
        max_percentage: int,
        force_large_change: bool,
    ) -> None:
        creates = [r for r in plan_rows if r["action"] == ACTION_CREATE]
        updates = [r for r in plan_rows if r["action"] in (ACTION_UPDATE, ACTION_CONVERT)]
        deletes = [r for r in plan_rows if r["action"] == ACTION_DELETE]

        if apply_mode and existing_child_count + len(creates) > MAX_CHILDREN_PER_PARENT:
            raise PreflightError(
                f"Planned result would have {existing_child_count + len(creates)} child tags "
                f"under {TARGET_PARENT_TAG_NAME!r}, exceeding the Qualys limit of "
                f"{MAX_CHILDREN_PER_PARENT}. Aborting before any mutation (no partial creation)."
            )

        if not apply_mode:
            return  # blast-radius gates only bind real writes

        total = len(creates) + len(updates) + (len(deletes) if allow_delete else 0)
        pct = (total / existing_child_count * 100) if existing_child_count else 0
        problems = []
        if len(creates) > max_create:
            problems.append(f"creates {len(creates)} > max {max_create}")
        if len(updates) > max_update:
            problems.append(f"updates {len(updates)} > max {max_update}")
        if allow_delete and len(deletes) > max_delete:
            problems.append(f"deletes {len(deletes)} > max {max_delete}")
        if total > max_total:
            problems.append(f"total changes {total} > max {max_total}")
        if existing_child_count and pct > max_percentage:
            problems.append(f"{pct:.1f}% of existing child tags would change > max {max_percentage}%")

        if problems and not force_large_change:
            raise PreflightError(
                "Blast-radius guardrail(s) tripped: " + "; ".join(problems) + ". "
                "Re-run with --force-large-change to override, after reviewing the dry-run report."
            )
        if problems:
            print("WARNING: blast-radius guardrail(s) overridden by --force-large-change: " + "; ".join(problems))


# ==========================================================================
# Backup
# ==========================================================================


class BackupManager:
    @staticmethod
    def write(touched_tags: list[QualysTag], timestamp: str, output_dir: str) -> tuple[str, str]:
        json_path = os.path.join(output_dir, f"backup_application_tags_{timestamp}.json")
        xlsx_path = os.path.join(output_dir, f"backup_application_tags_{timestamp}.xlsx")

        payload = [
            {
                "tag_id": t.tag_id,
                "tag_name": t.tag_name,
                "parent_tag_id": t.parent_tag_id,
                "rule_type": t.rule_type,
                "rule_text": t.rule_text,
                "color": t.color,
                "criticality": t.criticality,
                "description": t.description,
                "created": t.created,
                "modified": t.modified,
            }
            for t in touched_tags
        ]
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)

        wb = Workbook()
        ws = wb.active
        ws.title = "Backup"
        cols = list(payload[0].keys()) if payload else [
            "tag_id", "tag_name", "parent_tag_id", "rule_type", "rule_text",
            "color", "criticality", "description", "created", "modified",
        ]
        ws.append(cols)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for row in payload:
            ws.append([row.get(c, "") for c in cols])
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{ws.max_row}"
        wb.save(xlsx_path)

        print(f"Backup written: {json_path}")
        print(f"Backup written: {xlsx_path}")
        return json_path, xlsx_path


# ==========================================================================
# Execution
# ==========================================================================


class ChangeExecutor:
    def __init__(self, client: QualysTagClient):
        self.client = client

    def apply_update_or_convert(self, row: dict) -> None:
        tag: QualysTag = row["_tag"]
        tag_id = tag.tag_id

        live = self.client.fetch_by_id(tag_id)
        if live is None or live.tag_name != tag.tag_name:
            row["verification_status"] = "FAILED"
            row["error_message"] = "aborted: tag changed or vanished immediately before write"
            return

        fields = []
        new_ips = row.get("_new_ips")
        if new_ips is not None and not row.get("_color_only") and not row.get("_rename_only"):
            fields.append(f"<ruleType>{EXPECTED_RULE_TYPE}</ruleType>")
            fields.append(f"<ruleText>{xml_escape(','.join(new_ips))}</ruleText>")
        new_name = row.get("_new_name")
        if new_name:
            fields.append(f"<name>{xml_escape(new_name)}</name>")
        new_color = row.get("_new_color")
        if new_color:
            fields.append(f"<color>{xml_escape(new_color)}</color>")

        if not fields:
            row["verification_status"] = "NOT_APPLICABLE"
            return

        try:
            self.client.update(tag_id, "".join(fields), f"update tag {tag_id}")
        except RuntimeError as exc:
            row["verification_status"] = "FAILED"
            row["error_message"] = str(exc)
            return

        verify = self.client.fetch_by_id(tag_id)
        if verify is None:
            row["verification_status"] = "FAILED"
            row["error_message"] = "update call succeeded but tag not found on read-back"
            return

        if new_ips is not None and not row.get("_color_only") and not row.get("_rename_only"):
            stored_ips, _ = parse_qualys_rule_text(verify.rule_text)
            if set(stored_ips) != set(new_ips):
                row["verification_status"] = "FAILED"
                row["error_message"] = (
                    f"post-update verification failed: stored={stored_ips!r} expected={new_ips!r}"
                )
                return
            if verify.rule_type != EXPECTED_RULE_TYPE:
                row["verification_status"] = "FAILED"
                row["error_message"] = f"post-update verification failed: ruleType={verify.rule_type!r}"
                return

        if new_name and verify.tag_name != new_name:
            row["verification_status"] = "FAILED"
            row["error_message"] = f"post-update verification failed: name={verify.tag_name!r} expected={new_name!r}"
            return

        if new_color and normalize_color(verify.color) != normalize_color(new_color):
            row["verification_status"] = "FAILED"
            row["error_message"] = f"post-update verification failed: color={verify.color!r}"
            return

        row["verification_status"] = "VERIFIED"
        if new_name:
            row["qualys_tag_name"] = new_name

    def apply_create(self, row: dict, parent_tag_id: str) -> None:
        name = row["qualys_tag_name"]
        existing = self.client.fetch_by_name(name)
        if existing is not None:
            row["verification_status"] = "FAILED"
            row["error_message"] = f"aborted: tag {name!r} already exists (id {existing.tag_id})"
            return

        new_ips = row.get("_new_ips") or []
        parts = [
            f"<name>{xml_escape(name)}</name>",
            f"<parentTagId>{xml_escape(str(parent_tag_id))}</parentTagId>",
        ]
        if row.get("_create_static"):
            pass  # omit ruleType/ruleText entirely -> Qualys creates a STATIC tag
        else:
            parts.append(f"<ruleType>{EXPECTED_RULE_TYPE}</ruleType>")
            parts.append(f"<ruleText>{xml_escape(','.join(new_ips))}</ruleText>")
        color = row.get("_new_color")
        if color:
            parts.append(f"<color>{xml_escape(color)}</color>")
        description = row.get("_description")
        if description:
            parts.append(f"<description>{xml_escape(description)}</description>")

        try:
            self.client.create("".join(parts), f"create tag {name!r}")
        except RuntimeError as exc:
            row["verification_status"] = "FAILED"
            row["error_message"] = str(exc)
            return

        verify = self.client.fetch_by_name(name)
        if verify is None:
            row["verification_status"] = "FAILED"
            row["error_message"] = "create call succeeded but tag not found on read-back"
            return
        if verify.parent_tag_id != str(parent_tag_id):
            row["verification_status"] = "FAILED"
            row["error_message"] = f"post-create verification failed: parent={verify.parent_tag_id!r}"
            return
        if not row.get("_create_static"):
            stored_ips, _ = parse_qualys_rule_text(verify.rule_text)
            if set(stored_ips) != set(new_ips):
                row["verification_status"] = "FAILED"
                row["error_message"] = (
                    f"post-create verification failed: stored={stored_ips!r} expected={new_ips!r}"
                )
                return

        row["qualys_tag_id"] = verify.tag_id
        row["verification_status"] = "VERIFIED"

    def apply_delete(self, row: dict, deletable_names: set[str]) -> None:
        tag: QualysTag = row["_tag"]
        tag_id = tag.tag_id
        expected_name = tag.tag_name

        live = self.client.fetch_by_id(tag_id)
        if live is None:
            row["verification_status"] = "FAILED"
            row["error_message"] = "aborted: tag not found immediately before delete"
            return
        if live.tag_name != expected_name:
            row["verification_status"] = "FAILED"
            row["error_message"] = (
                f"aborted: id {tag_id} now carries name {live.tag_name!r}, expected {expected_name!r}"
            )
            return
        if expected_name not in deletable_names:
            row["verification_status"] = "FAILED"
            row["error_message"] = "aborted: tag is no longer an approved deletion candidate"
            return

        try:
            self.client.delete(tag_id, f"delete tag {tag_id}")
        except RuntimeError as exc:
            row["verification_status"] = "FAILED"
            row["error_message"] = str(exc)
            return

        if self.client.fetch_by_id(tag_id) is not None:
            row["verification_status"] = "FAILED"
            row["error_message"] = "delete call succeeded but tag still present on read-back"
            return

        row["verification_status"] = "VERIFIED"


# ==========================================================================
# Reporting
# ==========================================================================


class ReportGenerator:
    ACTION_EXPLANATIONS = {
        ACTION_CREATE: "No matching tag existed; created under the parent hierarchy.",
        ACTION_UPDATE: "Existing tag's ruleText/colour/name rewritten in place to match the CMDB.",
        ACTION_CONVERT: "Existing STATIC tag converted in place to a dynamic NETWORK_RANGE tag.",
        ACTION_DELETE: "Tag has no corresponding live CMDB application; removed (gated by --allow-delete).",
        ACTION_NO_CHANGE: "Already matches the CMDB, or is intentionally preserved untouched.",
        ACTION_SKIP_NO_IPS: "CMDB has no usable IPs for this application; no tag created.",
        ACTION_SKIP_UNSUPPORTED: "Tag's ruleType is never safely writable by this script (e.g. NAME_CONTAINS).",
        ACTION_SKIP_AMBIGUOUS: "Tag name does not follow the expected pattern; left alone for manual review.",
        ACTION_ERROR: "Failed validation, verification, or an unexpected API/tag state.",
    }

    @staticmethod
    def _clean(rows: list[dict]) -> list[dict]:
        return [{k: v for k, v in r.items() if not k.startswith("_")} for r in rows]

    @classmethod
    def write_csv(cls, rows: list[dict], path: str) -> None:
        clean = cls._clean(rows)
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=REPORT_COLUMNS)
            writer.writeheader()
            for r in clean:
                writer.writerow({col: r.get(col, "") for col in REPORT_COLUMNS})

    @classmethod
    def write_xlsx(cls, rows: list[dict], path: str, run_meta: dict, summary_counts: dict) -> None:
        clean = cls._clean(rows)
        wb = Workbook()
        wb.remove(wb.active)

        ws = wb.create_sheet("Summary")
        ws.append([f"[VFZ] Applications & Platforms tag reconciliation - {run_meta.get('Run mode', '')}"])
        ws["A1"].font = Font(bold=True, size=14)
        ws.append([])
        for key, value in run_meta.items():
            ws.append([key, value])
        ws.append([])
        ws.append(["Metric", "Value"])
        for cell in ws[ws.max_row]:
            cell.font = Font(bold=True)
        for key, value in summary_counts.items():
            ws.append([key, value])
        ws.append([])
        ws.append(["Action", "Rows", "What it means"])
        for cell in ws[ws.max_row]:
            cell.font = Font(bold=True)

        buckets: dict[str, list[dict]] = {}
        for r in clean:
            buckets.setdefault(r["action"], []).append(r)
        for action, bucket_rows in buckets.items():
            ws.append([action, len(bucket_rows), cls.ACTION_EXPLANATIONS.get(action, "")])

        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 10
        ws.column_dimensions["C"].width = 80

        for action, bucket_rows in buckets.items():
            cls._add_sheet(wb, action[:31], bucket_rows)

        cls._add_sheet(wb, "All Applications", clean)
        wb.save(path)

    @staticmethod
    def _add_sheet(wb, title, rows):
        ws = wb.create_sheet(title)
        ws.append(REPORT_COLUMNS)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for r in rows:
            ws.append([r.get(c, "") for c in REPORT_COLUMNS])
        wrap_idx = {REPORT_COLUMNS.index(c) + 1 for c in WRAP_COLUMNS}
        for row_num in range(2, ws.max_row + 1):
            for col_idx in wrap_idx:
                ws.cell(row=row_num, column=col_idx).alignment = Alignment(wrap_text=True, vertical="top")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(REPORT_COLUMNS))}{ws.max_row}"
        for idx, col in enumerate(REPORT_COLUMNS, start=1):
            ws.column_dimensions[get_column_letter(idx)].width = REPORT_COLUMN_WIDTHS.get(col, 16)


def _write_with_fallback(writer, path):
    """Write, falling back to a timestamped name if the file is locked (e.g.
    still open in Excel on Windows)."""
    try:
        writer(path)
        return path
    except PermissionError:
        stem, ext = os.path.splitext(path)
        stamp = datetime.datetime.now().strftime("%H%M%S")
        alt = f"{stem}_{stamp}{ext}"
        writer(alt)
        print(f"NOTE: {path} was locked; wrote {alt} instead.")
        return alt


# ==========================================================================
# CLI / main
# ==========================================================================


def parse_args():
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("--file", required=True, help="Path to the CMDB Excel export")

    mode = parser.add_mutually_exclusive_group()
    mode.add_argument(
        "--dry-run", action="store_true", help="Preview only, no Qualys writes (default)."
    )
    mode.add_argument(
        "--apply", action="store_true", help="Perform real, non-destructive Qualys writes."
    )
    parser.add_argument(
        "--allow-delete",
        action="store_true",
        help="Also permit deletions. Requires --apply. Ignored otherwise.",
    )
    parser.add_argument(
        "--manage-color",
        action="store_true",
        help="Reconcile existing tags' colour to DEFAULT_TAG_COLOR when it differs.",
    )
    parser.add_argument(
        "--create-static-for-no-ip",
        action="store_true",
        help="Create a placeholder STATIC tag for CMDB applications with no usable IPs "
        "(default: report MISSING_TAG_NO_IP_SCOPE and create nothing).",
    )
    parser.add_argument(
        "--rename-tags",
        action="store_true",
        help="Allow matching an existing tag to an application via a GAID recorded in its "
        "description when the exact name no longer matches, and rename it in place.",
    )
    parser.add_argument("--max-create", type=int, default=MAX_CREATE_CHANGES)
    parser.add_argument("--max-update", type=int, default=MAX_UPDATE_CHANGES)
    parser.add_argument("--max-delete", type=int, default=MAX_DELETE_CHANGES)
    parser.add_argument("--max-total", type=int, default=MAX_TOTAL_CHANGES)
    parser.add_argument("--max-percentage-changed", type=int, default=MAX_PERCENTAGE_CHANGED)
    parser.add_argument(
        "--force-large-change",
        action="store_true",
        help="Override the blast-radius guardrails above.",
    )
    parser.add_argument("--verbose", action="store_true")
    parser.add_argument("--output-dir", default=".", help="Directory for reports and backups.")
    return parser.parse_args()


def _print_plan(rows: list[dict]) -> None:
    for r in rows:
        action = r["action"]
        if action == ACTION_NO_CHANGE:
            continue
        print(f"  [{action}] {r['application'] or r['qualys_tag_name']} "
              f"(tag {r['qualys_tag_id'] or '(new)'}) -- {r['reason']}")
        if r.get("ips_added"):
            print(f"      + added:   {r['ips_added']}")
        if r.get("ips_removed"):
            print(f"      - removed: {r['ips_removed']}")
    print()


def _print_summary(rows: list[dict], applied: bool) -> None:
    counts: dict[str, int] = {}
    for r in rows:
        counts[r["action"]] = counts.get(r["action"], 0) + 1
    print("=== Summary ===")
    print(f"Mode: {'APPLY' if applied else 'DRY-RUN'}")
    for action in (
        ACTION_CREATE, ACTION_UPDATE, ACTION_CONVERT, ACTION_DELETE, ACTION_NO_CHANGE,
        ACTION_SKIP_NO_IPS, ACTION_SKIP_UNSUPPORTED, ACTION_SKIP_AMBIGUOUS, ACTION_ERROR,
    ):
        if counts.get(action):
            print(f"  {action:28s} {counts[action]}")
    if applied:
        failed = sum(1 for r in rows if r.get("verification_status") == "FAILED")
        print(f"  Failed verifications:       {failed}")


def _make_console_encoding_safe() -> None:
    """CMDB ASSET names have been observed to carry stray Unicode (e.g. a
    zero-width space from a copy/paste) that a Windows console's legacy
    codepage cannot render. Without this, a plain print() of such a value
    crashes the whole run with UnicodeEncodeError deep into printing the
    plan -- after CMDB parsing and the Qualys tag fetch already happened.
    Report files are unaffected (they are opened with encoding="utf-8"
    explicitly); this only hardens interactive console output.
    """
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            try:
                reconfigure(errors="replace")
            except (ValueError, OSError):
                pass


def main() -> int:
    _make_console_encoding_safe()
    args = parse_args()
    apply_mode = bool(args.apply)
    allow_delete = apply_mode and bool(args.allow_delete)
    if args.allow_delete and not apply_mode:
        print("NOTE: --allow-delete has no effect without --apply; ignoring.")

    os.makedirs(args.output_dir, exist_ok=True)
    run_timestamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%dT%H%M%SZ")

    if apply_mode:
        print("*" * 72)
        print("APPLY MODE - changes will be written to Qualys" + (", INCLUDING DELETIONS" if allow_delete else ""))
        print("Every tag that will be modified/converted/deleted is backed up first.")
        print("*" * 72)
        print()

    headers, rows, columns = CmdbReader(args.file).load()
    builder = DesiredStateBuilder(headers, rows, columns)
    applications = builder.build()
    PreflightValidator.validate_source(builder)

    print(f"Found {len(applications)} distinct application(s) (ASSET) in the CMDB.")
    with_ips = sum(1 for a in applications.values() if a.desired_ips)
    print(f"  {with_ips} have usable IPs after status filtering and network exclusion.")
    print()

    username, password = CredentialProvider.get()
    session = requests.Session()
    session.auth = (username, password)
    client = QualysTagClient(session)

    print(f"=== Resolving parent tag {TARGET_PARENT_TAG_NAME!r} ===")
    tag_model = QualysTagModel.resolve(client)
    print(f"Parent tag id: {tag_model.parent.tag_id}; {len(tag_model.children)} existing child tag(s).")
    print()

    planner = ChangePlanner(
        applications,
        tag_model,
        manage_color=args.manage_color,
        create_tags_without_ips=args.create_static_for_no_ip,
        rename_tags=args.rename_tags,
    )
    plan_rows = planner.plan()

    try:
        PreflightValidator.validate_plan(
            plan_rows,
            existing_child_count=len(tag_model.children),
            apply_mode=apply_mode,
            allow_delete=allow_delete,
            max_create=args.max_create,
            max_update=args.max_update,
            max_delete=args.max_delete,
            max_total=args.max_total,
            max_percentage=args.max_percentage_changed,
            force_large_change=args.force_large_change,
        )
    except PreflightError as exc:
        print(f"PREFLIGHT ABORT: {exc}")
        return 1

    if not allow_delete:
        for r in plan_rows:
            if r["action"] == ACTION_DELETE:
                r["verification_status"] = "NOT_APPLICABLE"
                r["error_message"] = "deletion candidate but --allow-delete not enabled"

    print("=== Plan ===")
    _print_plan(plan_rows)

    run_meta = {
        "Run mode": "APPLY" + (" (deletions enabled)" if allow_delete else " (deletions disabled)") if apply_mode else "DRY RUN",
        "Run at (UTC)": datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
        "Source file": os.path.basename(args.file),
        "Qualys pod": QUALYS_BASE_URL,
        "Parent tag": f"{TARGET_PARENT_TAG_NAME} (id {tag_model.parent.tag_id})",
        "Child tag name format": f"{CHILD_TAG_PREFIX}<ASSET>",
        "Resource status filter": ", ".join(sorted(INCLUDE_RESOURCE_STATUSES)),
        "Excluded IP blocks": ", ".join(EXCLUDED_IP_NETWORKS),
        "Manage colour": "yes" if args.manage_color else "no",
        "Create tags without IPs": "yes" if args.create_static_for_no_ip else "no",
        "Rename via GAID linkage": "yes" if args.rename_tags else "no",
    }
    summary_counts = {
        "Applications in source": len(applications),
        "Existing child tags in Qualys": len(tag_model.children),
        "Creates": sum(1 for r in plan_rows if r["action"] == ACTION_CREATE),
        "Updates": sum(1 for r in plan_rows if r["action"] == ACTION_UPDATE),
        "Conversions (STATIC->NETWORK_RANGE)": sum(1 for r in plan_rows if r["action"] == ACTION_CONVERT),
        "Deletes": sum(1 for r in plan_rows if r["action"] == ACTION_DELETE),
        "No change": sum(1 for r in plan_rows if r["action"] == ACTION_NO_CHANGE),
        "No usable IPs": sum(1 for r in plan_rows if r["action"] == ACTION_SKIP_NO_IPS),
        "Unsupported rule type": sum(1 for r in plan_rows if r["action"] == ACTION_SKIP_UNSUPPORTED),
        "Ambiguous source": sum(1 for r in plan_rows if r["action"] == ACTION_SKIP_AMBIGUOUS),
        "Validation/plan errors": sum(1 for r in plan_rows if r["action"] == ACTION_ERROR),
        "Total desired IPs": sum(r.get("new_ip_count") or 0 for r in plan_rows),
        "Total IPs added": sum(r.get("ips_added_count") or 0 for r in plan_rows),
        "Total IPs removed": sum(r.get("ips_removed_count") or 0 for r in plan_rows),
    }

    stem_suffix = "applied" if apply_mode else datetime.date.today().isoformat()
    csv_path = os.path.join(args.output_dir, f"application_tag_report_{run_timestamp}_{stem_suffix}.csv")
    xlsx_path = os.path.join(args.output_dir, f"application_tag_report_{run_timestamp}_{stem_suffix}.xlsx")

    if not apply_mode:
        _write_with_fallback(lambda p: ReportGenerator.write_csv(plan_rows, p), csv_path)
        _write_with_fallback(lambda p: ReportGenerator.write_xlsx(plan_rows, p, run_meta, summary_counts), xlsx_path)
        print(f"Wrote {csv_path}")
        print(f"Wrote {xlsx_path}")
        _print_summary(plan_rows, applied=False)
        print()
        print("DRY RUN complete -- no Qualys writes were made. Re-run with --apply to write.")
        return 0

    to_update = [r for r in plan_rows if r["action"] in (ACTION_UPDATE, ACTION_CONVERT)]
    to_create = [r for r in plan_rows if r["action"] == ACTION_CREATE]
    to_delete = [r for r in plan_rows if r["action"] == ACTION_DELETE] if allow_delete else []

    touched_tags = [r["_tag"] for r in to_update] + [r["_tag"] for r in to_delete]
    if touched_tags:
        BackupManager.write(touched_tags, run_timestamp, args.output_dir)

    executor = ChangeExecutor(client)
    print(f"=== Applying: {len(to_update)} update/convert, {len(to_create)} create, {len(to_delete)} delete ===")
    for r in to_update:
        executor.apply_update_or_convert(r)
    for r in to_create:
        executor.apply_create(r, tag_model.parent.tag_id)
    deletable_names = {r["qualys_tag_name"] for r in to_delete}
    for r in to_delete:
        executor.apply_delete(r, deletable_names)

    _write_with_fallback(lambda p: ReportGenerator.write_csv(plan_rows, p), csv_path)
    _write_with_fallback(lambda p: ReportGenerator.write_xlsx(plan_rows, p, run_meta, summary_counts), xlsx_path)
    print(f"Wrote {csv_path}")
    print(f"Wrote {xlsx_path}")
    _print_summary(plan_rows, applied=True)

    any_failed = any(r.get("verification_status") == "FAILED" for r in plan_rows)
    if any_failed:
        print()
        print("ERROR: one or more mutations FAILED verification -- see report for details.")
        return 1
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except PreflightError as exc:
        print(f"PREFLIGHT ABORT: {exc}", file=sys.stderr)
        sys.exit(1)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
