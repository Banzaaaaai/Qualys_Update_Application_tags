# Qualys Application Tag Updater

Reconciles the `[VFZ] <ASSET>` Asset Management tags under the
**`[VFZ] Applications & Platforms`** parent in a Qualys tenant (EU2 pod)
against a CMDB Excel export (`CMDB.xlsx`, maintained monthly by the Service
Platforms & E2E Solutions team per the VodafoneZiggo CSAM architecture
document), which is treated as authoritative for both *which* applications
have a tag and *what IP scope* each tag should carry.

This project is architecturally derived from
[`Qualys_Update_GAID_tags`](https://github.com/Banzaaaaai/Qualys_Update_GAID_tags)
— it reuses that project's proven retry/pagination/IP-normalization/
backup/verification machinery — but it is **not** a GAID-tag script and does
not modify `[VFZ] Global Application Inventory`. It is a separate, standalone
tool for a separate hierarchy.

> **Default run mode is DRY-RUN.** Nothing is written to Qualys unless you
> pass `--apply`, and no tag is ever deleted unless you pass **both**
> `--apply` and `--allow-delete`.

---

## Source-of-truth model: ASSET vs GAID

| Column | Role |
|---|---|
| **ASSET** | **Application identity.** The tag name is always `"[VFZ] " + ASSET`. |
| **GAID** | Supporting identifier. Retained in the tag description (`"{ASSET} (GAID: {gaid})"`) and in every report row, but never the primary match key and never part of the tag name. |
| **RVIT** | Application criticality indicator. Carried through to the report for context. |
| **ASSET STATUS** | Application lifecycle/governance context (e.g. "Decommissioned"). **Informational only** — never by itself drives IP scope or deletion. |
| **RESOURCE STATUS** | Resource lifecycle. **This** is what drives IP scope (see below) and contributes to the delete-by-status policy. |
| **IPADDRESS** | Technical IP scope for a resource row. |

The known CMDB export has a 1:1 relationship between GAID and ASSET. This is
validated in preflight — **every** GAID must map to exactly one ASSET and
vice versa. Any violation **fails closed**: a validation-error report is
written and no Qualys mutation is attempted, ever.

### Why this matters

Unlike the GAID-tag script (where the numeric GAID is a stable identity key),
here the ASSET string itself is the tag-name identity. If an ASSET's spelling
changes between CMDB exports, the old tag looks "absent from CMDB" (a delete
candidate) and a new tag looks "new" (a create). See **Known limitations**
below and the optional `--rename-tags` mitigation.

---

## The resource-status rule (critical)

**Only rows with `RESOURCE STATUS == "In Service"` contribute IP addresses**
to an application's desired scope. `ASSET STATUS` is never used for this.

This matters because the CMDB contains real cases where an application is
marked `Decommissioned` / `Out of Service` at the ASSET level while one or
more of its resources are still `In Service` — that application is alive and
its tag must keep working, or be created if missing.

---

## What the script does

| Situation | Action |
|---|---|
| ASSET in CMDB, has usable IPs, no matching tag | **CREATE** a `NETWORK_RANGE` tag under the parent |
| ASSET in CMDB, no usable IPs, no matching tag | Report `MISSING_TAG_NO_IP_SCOPE`; nothing created (unless `--create-static-for-no-ip`) |
| ASSET in CMDB, tag exists, `NETWORK_RANGE`, IPs differ | **UPDATE** `ruleText` — complete replacement, never a merge |
| ASSET in CMDB, tag exists, `NETWORK_RANGE`, IPs already match | `NO_CHANGE` |
| ASSET in CMDB, tag exists, `STATIC`, CMDB has usable IPs | **CONVERT** in place to `NETWORK_RANGE` — same tag id/name/parent |
| ASSET in CMDB, tag exists, `STATIC`, no usable IPs | Left **completely untouched**, reported `NO_USABLE_IPS` |
| ASSET in CMDB, tag exists, `NAME_CONTAINS` | **Never written to.** Reported `SKIPPED_UNSAFE_RULE_TYPE` |
| ASSET in CMDB, tag exists, unrecognized rule type | **ERROR** — fails safe at the tag level, never guessed at |
| Tag exists, ASSET absent from CMDB entirely | **DELETE** candidate (gated by `--apply --allow-delete`) |
| Tag exists, every CMDB row for that ASSET is `Out of Service` | **DELETE** candidate, same gating (see delete policy below) |
| Tag exists, ASSET present with any usable/live resource | Never a delete candidate, even with zero current IPs |

An existing tag is **always updated in place** — same id, name, and parent.
A normal update writes only the fields that actually changed: a colour-only
fix sends just `<color>`; a `STATIC → NETWORK_RANGE` conversion sends
`ruleType` + `ruleText` together; nothing else is ever touched.

### The one rule that matters most: never empty a tag's scope

An application with **zero usable IPs right now** is never treated as "this
application is gone." A `STATIC` tag with no IPs is left alone. A
`NETWORK_RANGE` tag is **never cleared** to an empty rule just because this
month's export has no `In Service` IP rows for it. Both cases are reported
`NO_USABLE_IPS`. Clearing scope is a far stronger, and much harder to notice,
action than updating it — this script never does it as a side effect of
filtering.

### Delete policy

A child tag is a deletion candidate under exactly two conditions:

1. Its ASSET is nowhere in the CMDB export, or
2. `ENABLE_DELETE_BY_ALL_RESOURCES_OOS` is on (default `True`) **and** every
   single CMDB row for that ASSET — including rows with no IP address at
   all — has `RESOURCE STATUS == "Out of Service"`.

A blank/unknown resource status on even one row blocks deletion. This is
deliberate: a live resource that simply carries no IP (a database, a load
balancer) must not cause its application's tag to be deleted.

Deletion additionally requires **both** `--apply` and `--allow-delete`. In
`--dry-run`, or in `--apply` without `--allow-delete`, delete candidates are
reported but nothing is removed.

---

## IP handling

- `IPADDRESS` cells may contain single IPv4 addresses, comma/semicolon/
  newline-separated lists, and `A-B` hyphen ranges, in any mixture, across
  however many rows an ASSET has.
- Every token is validated, expanded to an integer address representation,
  unioned, deduplicated, and recompacted into the minimal sorted set of
  single-IP / `A-B` entries — so `10.1.1.1,10.1.1.2` and `10.1.1.1-10.1.1.2`
  from different sources are recognized as identical scope and never produce
  a spurious diff.
- Range expansion is capped by `MAX_RANGE_SIZE` (default 1,000,000) so a
  malformed range (e.g. a typo'd `/8`) can't exhaust memory.
- Tokens that fail validation are never silently dropped — they mark that
  application's plan row `ERROR` with the offending tokens listed, and no
  write is attempted for it.
- `EXCLUDED_IP_NETWORKS` (default `169.254.0.0/16`, `192.168.0.0/16`) are
  stripped **only from the CMDB-derived desired set**, never from what
  Qualys currently holds. This is what lets an excluded address that is
  already stored in a tag show up correctly as a removal in the diff instead
  of being invisible on both sides.

---

## Safety controls

- **Preflight, fail-closed.** GAID↔ASSET 1:1 validation and blank-ASSET rows
  abort the entire run with a validation-error report before any Qualys call
  is made. See `PreflightValidator.validate_source`.
- **Parent resolved by exact name, never hardcoded.** The script aborts if
  zero or more than one tag named `[VFZ] Applications & Platforms` exists.
- **350-child-tag ceiling.** Before any create, the script checks
  `existing_children + planned_creates` against Qualys' documented 350
  children-per-parent limit and fails closed (no partial creation) if it
  would be exceeded.
- **Blast-radius guardrails** (`--apply` only): `MAX_CREATE_CHANGES`,
  `MAX_UPDATE_CHANGES`, `MAX_DELETE_CHANGES`, `MAX_TOTAL_CHANGES`, and
  `MAX_PERCENTAGE_CHANGED` (percentage of existing child tags touched) all
  abort the run before any mutation if tripped. Override with
  `--force-large-change` only after reviewing the dry-run report.
- **Backups before destructive writes.** Every tag about to be updated,
  converted, or deleted is captured — full field set — to
  `backup_application_tags_<UTC>.json` and `.xlsx` before the first write
  call. Backups are never overwritten (each run gets its own UTC timestamp).
- **Read-back verification on every mutation.** Create, update, convert, and
  delete are each followed by re-fetching the tag from Qualys and checking
  it matches what was intended (id, name, parent, ruleType, ruleText,
  colour as applicable). A failed verification marks that row `FAILED` and
  the process exits non-zero — partial failures are never hidden.
- **Protected tags.** Anything matching `PROTECTED_TAG_PATTERNS` (by default
  just the parent tag's own name) is never a deletion candidate.
- **NAME_CONTAINS and unrecognized rule types are never guessed at.** They
  are reported and left completely alone.

---

## Rollback considerations

- The JSON/XLSX backups capture every field this script can read back for a
  tag (`ruleType`, `ruleText`, `color`, `description`, `criticality`, etc.),
  so a bad update or conversion can be manually reverted from the backup.
- **Deletion is irreversible** through this tool. It also detaches the tag
  from any Qualys access scope or asset-group logic bound to that tag's
  identity, which the backup cannot restore. Treat `--allow-delete` runs
  with the same care as any other irreversible production change, and always
  review the immediately preceding `--dry-run` report first.
- Because tag identity here is the ASSET name string (not a stable numeric
  id), a create+delete pair produced by an ASSET rename is **not** a like-for
  -like "restore" from backup — the backup restores the *old* tag's fields,
  it does not fix the identity mismatch. See `--rename-tags` below.

---

## Reports

Every run (dry-run or apply) produces both a CSV and an XLSX report,
timestamped and never overwritten:

```
application_tag_report_<UTC>_<dry-date-or-applied>.csv
application_tag_report_<UTC>_<dry-date-or-applied>.xlsx
```

The report is **application-oriented** (one row per ASSET, plus one row per
orphaned Qualys tag), not tag-oriented — the primary column is `application`,
with `gaid`, `rvit`, `asset_status`, and `resource_status_summary` alongside
it for context. Full column list:

```
action, application, gaid, rvit, asset_status, resource_status_summary,
qualys_tag_id, qualys_tag_name, qualys_rule_type, desired_rule_type,
old_ip_count, new_ip_count, ips_added_count, ips_removed_count,
ips_excluded_count, old_ip_summary, new_ip_summary, ips_added, ips_removed,
reason, verification_status, error_message
```

`action` is one of: `CREATE`, `UPDATE`, `CONVERT_STATIC_TO_DYNAMIC`,
`DELETE`, `NO_CHANGE`, `SKIP_NO_USABLE_IPS`, `SKIP_UNSUPPORTED_RULE_TYPE`,
`SKIP_AMBIGUOUS_SOURCE`, `ERROR`.

The XLSX report has a **Summary** sheet (run metadata + counts: source
application count, existing Qualys child-tag count, creates/updates/
conversions/deletes/skips, no-IP count, unsupported-rule count, validation
error count, total IPs desired/added/removed), one sheet per non-empty
action bucket, and an **All Applications** sheet with every row.

---

## Installation

Requires Python 3.11+.

```bash
pip install requests openpyxl
```

## Credentials

- Never passed on the command line, never hardcoded.
- Read from `QUALYS_USERNAME` / `QUALYS_PASSWORD` environment variables.
- Falling back to `qualys_creds.txt` next to the script:
  `user:<tab>your_username` / `pass:<tab>your_password` (one entry per line).
- `qualys_creds.txt` and the CMDB export (`asset-resource-owner.xlsx`) are
  git-ignored — never commit them.

## Usage

```bash
# Preview only (default) -- no Qualys writes of any kind
python update_application_tags.py --file CMDB.xlsx

# Apply non-destructive changes: creates, updates, static->dynamic conversions
python update_application_tags.py --file CMDB.xlsx --apply

# Apply everything, including deletions of tags absent from the CMDB
python update_application_tags.py --file CMDB.xlsx --apply --allow-delete
```

Useful options:

| Flag | Effect |
|---|---|
| `--file PATH` | CMDB Excel export (required) |
| `--dry-run` | Explicit no-op preview (this is also the default with no mode flag) |
| `--apply` | Perform real, non-destructive Qualys writes |
| `--allow-delete` | Also permit deletions (requires `--apply`) |
| `--manage-color` | Reconcile existing tags' colour to `DEFAULT_TAG_COLOR` (`#0000FF`) when it differs. New tags always get this colour regardless of this flag. |
| `--create-static-for-no-ip` | Create a placeholder `STATIC` tag for a CMDB application with no usable IPs, instead of reporting `MISSING_TAG_NO_IP_SCOPE` and creating nothing (default) |
| `--rename-tags` | Match an existing tag to an application via the GAID recorded in its description when the exact name no longer matches, and rename the tag in place instead of create+delete |
| `--max-create` / `--max-update` / `--max-delete` / `--max-total` / `--max-percentage-changed` | Override the blast-radius guardrails |
| `--force-large-change` | Override a tripped blast-radius guardrail (child-count ceiling is never overridable) |
| `--output-dir DIR` | Where reports and backups are written (default: current directory) |
| `--verbose` | More detail in console output |

---

## Known limitations

- **ASSET-name identity and renames.** Because the tag name is derived
  directly from the ASSET string (per the task's requirement that ASSET be
  the primary identity, not GAID), a spelling change in the CMDB's ASSET
  column between runs is indistinguishable from "old application removed,
  new application added" unless `--rename-tags` is used, and even then only
  when the tag's description still carries a recoverable `(GAID: n)` marker
  from a prior run of this script (pre-existing tags created before this
  tool was introduced will not have that marker).
- **"Is a parent tag" is inferred, not directly queried.** The Qualys QPS
  tag API does not expose a first-class "this tag has/permits children"
  flag; the parent is verified by exact-name uniqueness and by successfully
  listing its children, which is a practical rather than a structural
  guarantee.
- **Tag-name length ceiling (256 chars) is an assumed conservative limit**,
  not a documented Qualys constant; it exists so a pathological ASSET string
  fails at plan time with a clear message instead of an opaque API error.
- **Shared IPs across applications are not flagged by default.** The spec
  explicitly treats IP overlap between applications as legitimate, not an
  error; a dedicated "shared IP" forensic report was intentionally left out
  of this initial version to keep scope focused — the report's
  `application`-oriented rows already make it possible to grep the same IP
  across multiple rows manually.

---

## Testing

```bash
python -m unittest discover -s tests -v
```

All tests are pure/offline — the Qualys API client is never invoked. Coverage
includes: GAID/ASSET 1:1 validation, blank-ASSET detection, IP parsing and
range expansion/compaction, excluded-network filtering, resource-status
filtering, all-resources-out-of-service detection, no-usable-IP safety
behavior, `NETWORK_RANGE` diffing, `STATIC → NETWORK_RANGE` conversion,
`NAME_CONTAINS` skip behavior, create/delete gating, and blast-radius
threshold enforcement.
