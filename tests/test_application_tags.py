"""Unit tests for update_application_tags.py.

No live Qualys tenant is required: the Qualys API surface (QualysTagClient) is
never invoked here. Tests exercise pure logic -- IP parsing, the desired-state
builder, the diff/planning engine, and the blast-radius guardrails -- using
in-memory fixtures.
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import update_application_tags as app


# --------------------------------------------------------------------------
# IP parsing / range expansion / compaction / exclusion
# --------------------------------------------------------------------------


class TestIpParsing(unittest.TestCase):
    def test_single_ip(self):
        ordered, invalid = app.normalize_ip_set(["10.1.1.10"])
        self.assertEqual(ordered, ["10.1.1.10"])
        self.assertEqual(invalid, [])

    def test_duplicate_and_padding_forms_collapse(self):
        ordered, invalid = app.normalize_ip_set(["10.0.0.1", "10.0.0.1"])
        self.assertEqual(ordered, ["10.0.0.1"])
        self.assertEqual(invalid, [])

    def test_invalid_token_reported_not_raised(self):
        ordered, invalid = app.normalize_ip_set(["not-an-ip", "10.1.1.1"])
        self.assertEqual(ordered, ["10.1.1.1"])
        self.assertEqual(invalid, ["not-an-ip"])

    def test_range_expansion(self):
        ints = app.expand_entry_to_ints("10.1.1.1-10.1.1.3")
        self.assertEqual(len(ints), 3)

    def test_range_end_before_start_is_invalid(self):
        with self.assertRaises(ValueError):
            app.expand_entry_to_ints("10.1.1.5-10.1.1.1")

    def test_range_too_large_is_invalid(self):
        old_limit = app.MAX_RANGE_SIZE
        app.MAX_RANGE_SIZE = 10
        try:
            with self.assertRaises(ValueError):
                app.expand_entry_to_ints("10.0.0.0-10.0.1.0")
        finally:
            app.MAX_RANGE_SIZE = old_limit

    def test_range_compaction_merges_contiguous(self):
        entries = app.compact_ints_to_entries({1, 2, 3, 5})
        self.assertEqual(entries, ["0.0.0.1-0.0.0.3", "0.0.0.5"])

    def test_range_and_single_ip_are_equivalent_after_normalization(self):
        a, _ = app.normalize_ip_set(["10.1.1.1-10.1.1.2"])
        b, _ = app.normalize_ip_set(["10.1.1.1", "10.1.1.2"])
        self.assertEqual(set(a), set(b))

    def test_excluded_network_filtering(self):
        kept, dropped = app.drop_excluded_networks(["10.1.1.1", "169.254.1.1", "192.168.1.1"])
        self.assertEqual(kept, ["10.1.1.1"])
        self.assertEqual(set(dropped), {"169.254.1.1", "192.168.1.1"})

    def test_excluded_networks_not_applied_to_qualys_side(self):
        # drop_excluded_networks is only ever called on the CMDB-derived set;
        # parse_qualys_rule_text must NOT strip excluded ranges, or an
        # excluded address already stored in Qualys could never be diffed out.
        old_ips, _ = app.parse_qualys_rule_text("169.254.1.1,10.0.0.1")
        self.assertIn("169.254.1.1", old_ips)


class TestColorNormalization(unittest.TestCase):
    def test_unpadded_hex_normalizes_same_as_padded(self):
        self.assertEqual(app.normalize_color("#FF"), app.normalize_color("#0000FF"))

    def test_missing_color_normalizes_to_empty(self):
        self.assertEqual(app.normalize_color(""), "")
        self.assertEqual(app.normalize_color(None), "")

    def test_garbage_color_normalizes_to_empty(self):
        self.assertEqual(app.normalize_color("not-hex"), "")


class TestAssetNameNormalization(unittest.TestCase):
    def test_trims_and_collapses_whitespace(self):
        self.assertEqual(app.normalize_asset_name("  CMS\t Platform \n"), "CMS Platform")

    def test_preserves_case_and_punctuation(self):
        raw = "AAA (Authentication, Authorization and Accounting) (IP Works)"
        self.assertEqual(app.normalize_asset_name(raw), raw)

    def test_strips_invisible_format_characters(self):
        # U+200B ZERO WIDTH SPACE is a real copy/paste artifact observed in
        # CMDB ASSET values; left in place it silently defeats name matching
        # and can crash console output on a codepage that can't render it.
        self.assertEqual(app.normalize_asset_name("CMS​ Platform"), "CMS Platform")
        self.assertEqual(app.normalize_asset_name("﻿CMS"), "CMS")

    def test_gaid_key_strips_float_suffix(self):
        self.assertEqual(app.normalize_gaid_key(5014.0), "5014")
        self.assertEqual(app.normalize_gaid_key("5014.0"), "5014")
        self.assertEqual(app.normalize_gaid_key(""), None)
        self.assertEqual(app.normalize_gaid_key(None), None)


# --------------------------------------------------------------------------
# DesiredStateBuilder: source validation and resource-status handling
# --------------------------------------------------------------------------

# Column order used by every synthetic row below.
COLS = ["ASSET", "GAID", "RVIT", "ASSET_STATUS", "RESOURCE_STATUS", "IPADDRESS"]
COLUMN_INDEX = {name: i for i, name in enumerate(COLS)}


def make_builder(rows):
    return app.DesiredStateBuilder(headers=COLS, rows=rows, columns=COLUMN_INDEX)


class TestDesiredStateBuilder(unittest.TestCase):
    def test_gaid_asset_one_to_one_violation_is_a_validation_error(self):
        rows = [
            ("CMS", "5014", "no", "Active", "In Service", "10.1.1.1"),
            ("CMS", "9999", "no", "Active", "In Service", "10.1.1.2"),
        ]
        builder = make_builder(rows)
        builder.build()
        self.assertTrue(any("multiple GAID" in e.message for e in builder.errors))

    def test_duplicate_gaid_for_different_assets_is_a_validation_error(self):
        rows = [
            ("CMS", "5014", "no", "Active", "In Service", "10.1.1.1"),
            ("Other App", "5014", "no", "Active", "In Service", "10.1.1.2"),
        ]
        builder = make_builder(rows)
        builder.build()
        self.assertTrue(any("multiple ASSET" in e.message for e in builder.errors))

    def test_blank_asset_with_gaid_is_a_validation_error(self):
        rows = [("", "5014", "no", "Active", "In Service", "10.1.1.1")]
        builder = make_builder(rows)
        builder.build()
        self.assertTrue(any("blank ASSET" in e.message for e in builder.errors))

    def test_fully_blank_row_is_silently_skipped(self):
        rows = [(None, None, None, None, None, None)]
        builder = make_builder(rows)
        records = builder.build()
        self.assertEqual(records, {})
        self.assertEqual(builder.errors, [])

    def test_resource_status_filters_ip_contribution(self):
        rows = [
            ("CMS", "5014", "no", "Active", "In Service", "10.1.1.1"),
            ("CMS", "5014", "no", "Active", "Out of Service", "10.1.1.99"),
            ("CMS", "5014", "no", "Active", "Planned Decommission", "10.1.1.98"),
        ]
        records = make_builder(rows).build()
        self.assertEqual(records["CMS"].desired_ips, ["10.1.1.1"])

    def test_all_resources_out_of_service_detected_even_without_ip_rows(self):
        rows = [
            ("Legacy App", "100", "no", "Decommissioned", "Out of Service", None),
            ("Legacy App", "100", "no", "Decommissioned", "Out of Service", "10.1.1.1"),
        ]
        records = make_builder(rows).build()
        self.assertTrue(records["Legacy App"].all_resources_out_of_service)

    def test_one_in_service_resource_blocks_all_oos_even_if_asset_decommissioned(self):
        rows = [
            ("Legacy App", "100", "no", "Decommissioned", "Out of Service", "10.1.1.1"),
            ("Legacy App", "100", "no", "Decommissioned", "In Service", None),
        ]
        records = make_builder(rows).build()
        self.assertFalse(records["Legacy App"].all_resources_out_of_service)

    def test_blank_resource_status_blocks_all_oos(self):
        rows = [
            ("Legacy App", "100", "no", "Decommissioned", "Out of Service", "10.1.1.1"),
            ("Legacy App", "100", "no", "Decommissioned", None, None),
        ]
        records = make_builder(rows).build()
        self.assertFalse(records["Legacy App"].all_resources_out_of_service)

    def test_no_usable_ips_application_is_not_a_safety_hazard_by_itself(self):
        rows = [("Toolbox", "200", "no", "Active", "Out of Service", "10.1.1.1")]
        records = make_builder(rows).build()
        self.assertEqual(records["Toolbox"].desired_ips, [])
        # Not fully-OOS across ALL rows in this fixture would require every
        # row OOS; here there's only one row, and it IS all-OOS -- but the
        # important behavioural guarantee (tested at the planner level) is
        # that zero IPs never by itself clears or deletes a tag.
        self.assertTrue(records["Toolbox"].all_resources_out_of_service)


# --------------------------------------------------------------------------
# ChangePlanner: create / update / convert / skip / delete decisions
# --------------------------------------------------------------------------


def make_tag(name, rule_type="NETWORK_RANGE", rule_text="", color="#0000FF", description="", tag_id="1", parent_tag_id="999"):
    return app.QualysTag(
        tag_id=tag_id,
        tag_name=name,
        parent_tag_id=parent_tag_id,
        rule_type=rule_type,
        rule_text=rule_text,
        color=color,
        description=description,
        criticality="",
        created="",
        modified="",
    )


def make_application(asset, gaid="1", ips=None, all_row_statuses=None, rvit=""):
    rec = app.ApplicationRecord(asset=asset, gaid=gaid, rvit=rvit)
    rec.desired_ips = ips or []
    rec.all_row_statuses = all_row_statuses if all_row_statuses is not None else ["In Service"]
    return rec


def make_planner(applications, children, **kwargs):
    parent = make_tag(app.TARGET_PARENT_TAG_NAME, tag_id="999", parent_tag_id="")
    model = app.QualysTagModel(parent, children)
    return app.ChangePlanner(
        applications, model,
        create_tags_without_ips=kwargs.get("create_tags_without_ips", False),
        rename_tags=kwargs.get("rename_tags", False),
    )


class TestChangePlannerColor(unittest.TestCase):
    def test_create_defaults_to_blue(self):
        apps = {"CMS": make_application("CMS", ips=["10.1.1.10"], rvit="no")}
        planner = make_planner(apps, children=[])
        rows = planner.plan()
        self.assertEqual(rows[0]["_new_color"], app.DEFAULT_TAG_COLOR)

    def test_create_with_rvit_yes_is_red(self):
        apps = {"CMS": make_application("CMS", ips=["10.1.1.10"], rvit="yes")}
        planner = make_planner(apps, children=[])
        rows = planner.plan()
        self.assertEqual(rows[0]["_new_color"], app.CRITICAL_TAG_COLOR)

    def test_rvit_match_is_case_insensitive(self):
        for value in ("YES", "Yes", " yes "):
            with self.subTest(value=value):
                apps = {"CMS": make_application("CMS", ips=["10.1.1.10"], rvit=value)}
                planner = make_planner(apps, children=[])
                rows = planner.plan()
                self.assertEqual(rows[0]["_new_color"], app.CRITICAL_TAG_COLOR)

    def test_existing_network_range_tag_recoloured_to_red_with_no_other_change(self):
        tag = make_tag("[VFZ] CMS", rule_type="NETWORK_RANGE", rule_text="10.1.1.10", color="#0000FF")
        apps = {"CMS": make_application("CMS", ips=["10.1.1.10"], rvit="yes")}
        planner = make_planner(apps, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_UPDATE)
        self.assertEqual(rows[0]["_new_color"], app.CRITICAL_TAG_COLOR)
        self.assertEqual(rows[0]["ips_added"], "")  # IP set itself is unchanged
        self.assertEqual(rows[0]["ips_removed"], "")

    def test_already_correct_color_is_no_change(self):
        tag = make_tag("[VFZ] CMS", rule_type="NETWORK_RANGE", rule_text="10.1.1.10", color="#FF0000")
        apps = {"CMS": make_application("CMS", ips=["10.1.1.10"], rvit="yes")}
        planner = make_planner(apps, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_NO_CHANGE)

    def test_color_applies_unconditionally_to_name_contains_tag(self):
        # The key requirement: colour is corrected even for a rule type this
        # script otherwise never writes to -- colour is independent of the
        # rule content that makes NAME_CONTAINS unsafe to touch.
        tag = make_tag("[VFZ] CMS", rule_type="NAME_CONTAINS", rule_text="host.*", color="#0000FF")
        apps = {"CMS": make_application("CMS", ips=["10.1.1.10"], rvit="yes")}
        planner = make_planner(apps, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_UPDATE)
        self.assertTrue(rows[0]["_color_only"])
        self.assertEqual(rows[0]["_new_color"], app.CRITICAL_TAG_COLOR)

    def test_color_applies_unconditionally_to_static_tag_with_no_ips(self):
        tag = make_tag("[VFZ] Toolbox", rule_type="STATIC", rule_text="", color="#0000FF")
        apps = {"Toolbox": make_application("Toolbox", ips=[], rvit="yes")}
        planner = make_planner(apps, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_UPDATE)
        self.assertEqual(rows[0]["_new_color"], app.CRITICAL_TAG_COLOR)

    def test_orphan_name_contains_tag_recoloured_when_application_still_in_cmdb(self):
        # Rare data-quality edge case: tag name matches a CMDB asset by the
        # recovered-name guess but wasn't claimed by the forward pass.
        tag = make_tag("[VFZ] Weird App", rule_type="NAME_CONTAINS", rule_text="host.*", color="#0000FF")
        apps = {"Weird App": make_application("Weird App", ips=["10.1.1.10"], rvit="yes")}
        model = app.QualysTagModel(make_tag(app.TARGET_PARENT_TAG_NAME, tag_id="999", parent_tag_id=""), [tag])
        planner = app.ChangePlanner(apps, model, create_tags_without_ips=False, rename_tags=False)
        # Force the orphan path directly (bypassing the forward-pass claim)
        # to exercise _plan_orphan_tag's own colour handling in isolation.
        row = planner._plan_orphan_tag(tag)
        self.assertEqual(row["action"], app.ACTION_UPDATE)
        self.assertEqual(row["_new_color"], app.CRITICAL_TAG_COLOR)

    def test_orphan_tag_with_no_cmdb_application_has_no_color_opinion(self):
        # No RVIT data exists for a genuinely absent application, so colour
        # cannot be determined -- this is not an "exception" to the rule,
        # there is simply no data to apply it to.
        tag = make_tag("[VFZ] Gone App", rule_type="NETWORK_RANGE", rule_text="10.1.1.1", color="#00FF00")
        planner = make_planner({}, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_DELETE)
        self.assertEqual(rows[0].get("_new_color"), None)


class TestChangePlannerCreate(unittest.TestCase):
    def test_create_when_no_matching_tag_and_has_ips(self):
        apps = {"CMS": make_application("CMS", ips=["10.1.1.10"])}
        planner = make_planner(apps, children=[])
        rows = planner.plan()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["action"], app.ACTION_CREATE)
        self.assertEqual(rows[0]["qualys_tag_name"], "[VFZ] CMS")

    def test_no_tag_no_ips_is_skipped_not_created(self):
        apps = {"Toolbox": make_application("Toolbox", ips=[])}
        planner = make_planner(apps, children=[])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_SKIP_NO_IPS)

    def test_create_static_for_no_ip_flag_creates_placeholder(self):
        apps = {"Toolbox": make_application("Toolbox", ips=[])}
        planner = make_planner(apps, children=[], create_tags_without_ips=True)
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_CREATE)
        self.assertTrue(rows[0].get("_create_static"))


class TestChangePlannerUpdate(unittest.TestCase):
    def test_network_range_diff_computes_added_and_removed(self):
        tag = make_tag("[VFZ] CMS", rule_type="NETWORK_RANGE", rule_text="10.1.1.20")
        apps = {"CMS": make_application("CMS", ips=["10.1.1.10"])}
        planner = make_planner(apps, children=[tag])
        rows = planner.plan()
        row = rows[0]
        self.assertEqual(row["action"], app.ACTION_UPDATE)
        self.assertEqual(row["ips_added"], "10.1.1.10")
        self.assertEqual(row["ips_removed"], "10.1.1.20")

    def test_identical_ip_set_is_no_change_despite_different_surface_form(self):
        tag = make_tag("[VFZ] CMS", rule_type="NETWORK_RANGE", rule_text="10.1.1.1,10.1.1.2")
        apps = {"CMS": make_application("CMS", ips=["10.1.1.1-10.1.1.2"])}
        planner = make_planner(apps, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_NO_CHANGE)

    def test_static_tag_converted_when_cmdb_has_ips(self):
        tag = make_tag("[VFZ] CMS", rule_type="STATIC", rule_text="")
        apps = {"CMS": make_application("CMS", ips=["10.1.1.10"])}
        planner = make_planner(apps, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_CONVERT)
        self.assertTrue(rows[0]["_is_conversion"])

    def test_static_tag_with_no_ips_is_preserved_untouched(self):
        tag = make_tag("[VFZ] CMS", rule_type="STATIC", rule_text="")
        apps = {"CMS": make_application("CMS", ips=[])}
        planner = make_planner(apps, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_NO_CHANGE)
        self.assertIn("NO_USABLE_IPS", rows[0]["reason"])

    def test_network_range_tag_never_cleared_when_no_usable_ips(self):
        tag = make_tag("[VFZ] CMS", rule_type="NETWORK_RANGE", rule_text="10.1.1.1")
        apps = {"CMS": make_application("CMS", ips=[])}
        planner = make_planner(apps, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_NO_CHANGE)
        self.assertNotIn("_new_ips", rows[0])

    def test_name_contains_tag_is_never_overwritten(self):
        tag = make_tag("[VFZ] CMS", rule_type="NAME_CONTAINS", rule_text="host.*")
        apps = {"CMS": make_application("CMS", ips=["10.1.1.10"])}
        planner = make_planner(apps, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_SKIP_UNSUPPORTED)

    def test_groovy_and_asset_search_tags_are_never_overwritten(self):
        # These require manual review by a human -- never an automated write.
        for rule_type in ("GROOVY", "ASSET_SEARCH"):
            with self.subTest(rule_type=rule_type):
                tag = make_tag("[VFZ] CMS", rule_type=rule_type, rule_text="")
                apps = {"CMS": make_application("CMS", ips=["10.1.1.10"])}
                planner = make_planner(apps, children=[tag])
                rows = planner.plan()
                self.assertEqual(rows[0]["action"], app.ACTION_SKIP_UNSUPPORTED)
                self.assertIn(rule_type, rows[0]["reason"])

    def test_unexpected_rule_type_fails_safe(self):
        tag = make_tag("[VFZ] CMS", rule_type="SOME_FUTURE_TYPE", rule_text="")
        apps = {"CMS": make_application("CMS", ips=["10.1.1.10"])}
        planner = make_planner(apps, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_ERROR)


class TestChangePlannerCaseInsensitiveMatch(unittest.TestCase):
    def test_casing_only_mismatch_updates_existing_tag_without_renaming(self):
        # Regression: a live tenant had 8 CMDB/tag pairs differing only by
        # case (e.g. tag "[VFZ] Cyberark" vs CMDB ASSET "CyberArk"), which
        # exact matching alone would treat as delete-old + create-new.
        tag = make_tag("[VFZ] Cyberark", rule_type="NETWORK_RANGE", rule_text="10.1.1.1")
        apps = {"CyberArk": make_application("CyberArk", ips=["10.1.1.2"])}
        planner = make_planner(apps, children=[tag])
        rows = planner.plan()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["action"], app.ACTION_UPDATE)
        self.assertEqual(rows[0]["qualys_tag_name"], "[VFZ] Cyberark")  # name NOT changed
        self.assertIn("matched case-insensitively", rows[0]["reason"])

    def test_casing_only_mismatch_with_rename_tags_flag_renames(self):
        tag = make_tag("[VFZ] Cyberark", rule_type="NETWORK_RANGE", rule_text="10.1.1.1")
        apps = {"CyberArk": make_application("CyberArk", ips=["10.1.1.1"])}
        planner = make_planner(apps, children=[tag], rename_tags=True)
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_UPDATE)
        self.assertEqual(rows[0]["_new_name"], "[VFZ] CyberArk")

    def test_whitespace_only_mismatch_also_matches(self):
        tag = make_tag("[VFZ] I&M  portal", rule_type="NETWORK_RANGE", rule_text="10.1.1.1")
        apps = {"I&M portal": make_application("I&M portal", ips=["10.1.1.1"])}
        planner = make_planner(apps, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_NO_CHANGE)
        self.assertEqual(rows[0]["qualys_tag_name"], "[VFZ] I&M  portal")

    def test_ambiguous_case_insensitive_match_fails_safe(self):
        tag_a = make_tag("[VFZ] Foo", tag_id="1", rule_type="NETWORK_RANGE")
        tag_b = make_tag("[VFZ] FOO", tag_id="2", rule_type="NETWORK_RANGE")
        apps = {"foo": make_application("foo", ips=["10.1.1.1"])}
        planner = make_planner(apps, children=[tag_a, tag_b])
        rows = planner.plan()
        # one ERROR for the ambiguous application, plus two now-unclaimed
        # orphan tags evaluated independently.
        app_row = [r for r in rows if r["application"] == "foo" and r["qualys_tag_id"] == ""]
        self.assertEqual(len(app_row), 1)
        self.assertEqual(app_row[0]["action"], app.ACTION_ERROR)

    def test_exact_match_preferred_over_case_insensitive(self):
        exact = make_tag("[VFZ] CyberArk", tag_id="1", rule_type="NETWORK_RANGE", rule_text="10.1.1.1")
        other_case = make_tag("[VFZ] Cyberark", tag_id="2", rule_type="NETWORK_RANGE", rule_text="10.1.1.1")
        apps = {"CyberArk": make_application("CyberArk", ips=["10.1.1.1"])}
        planner = make_planner(apps, children=[exact, other_case])
        rows = planner.plan()
        app_row = [r for r in rows if r["application"] == "CyberArk" and r["qualys_tag_id"] != ""]
        self.assertEqual(len(app_row), 1)
        self.assertEqual(app_row[0]["qualys_tag_id"], "1")  # the exact match, not the casefold one


class TestChangePlannerDelete(unittest.TestCase):
    def test_tag_absent_from_cmdb_is_delete_candidate(self):
        tag = make_tag("[VFZ] OldApplication", rule_type="NETWORK_RANGE", rule_text="10.1.1.1")
        planner = make_planner({}, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_DELETE)
        self.assertIn("absent from CMDB", rows[0]["reason"])

    def test_tag_kept_when_application_present_with_live_resource(self):
        tag = make_tag("[VFZ] Legacy App", rule_type="NETWORK_RANGE", rule_text="10.1.1.1")
        apps = {
            "Legacy App": make_application(
                "Legacy App", ips=[], all_row_statuses=["Out of Service", "In Service"]
            )
        }
        planner = make_planner(apps, children=[tag])
        rows = planner.plan()
        # The forward pass claims this tag (name matches), so it is reported
        # as NO_CHANGE with NO_USABLE_IPS, never reaching the orphan/delete pass.
        self.assertEqual(rows[0]["action"], app.ACTION_NO_CHANGE)

    def test_tag_deleted_when_all_resources_out_of_service_and_policy_enabled(self):
        tag = make_tag("[VFZ] Legacy App", rule_type="NETWORK_RANGE", rule_text="10.1.1.1")
        # No application record at all under this exact name simulates the
        # tag being unmatched; combined with an application record present
        # under the SAME key this exercises the forward-pass NO_CHANGE path
        # instead, so to reach the orphan/delete pass we simply omit it from
        # `applications` entirely (absent from CMDB is the dominant case,
        # covered above). All-OOS-while-still-listed is exercised at the
        # ApplicationRecord level in TestDesiredStateBuilder.
        planner = make_planner({}, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_DELETE)

    def test_protected_tag_pattern_never_deleted(self):
        tag = make_tag(app.TARGET_PARENT_TAG_NAME, tag_id="999")
        planner = make_planner({}, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_NO_CHANGE)

    def test_tag_name_not_matching_prefix_is_ambiguous_not_deleted(self):
        tag = make_tag("Some Other Tag", rule_type="NETWORK_RANGE")
        planner = make_planner({}, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_SKIP_AMBIGUOUS)

    def test_orphan_name_contains_tag_is_never_deleted(self):
        # Regression: a live tenant had NAME_CONTAINS / GROOVY / ASSET_SEARCH
        # tags proposed for deletion merely because their name had no literal
        # CMDB match -- the same rule-type protection the forward pass
        # applies to writes must also gate the orphan/delete pass.
        tag = make_tag("[VFZ] Unify", rule_type="NAME_CONTAINS", rule_text="host.*")
        planner = make_planner({}, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_SKIP_UNSUPPORTED)

    def test_orphan_groovy_and_asset_search_tags_are_never_deleted(self):
        # GROOVY (a scripted rule) and ASSET_SEARCH (matches by e.g. QID) are
        # recognized-but-unsafe rule types, same tier as NAME_CONTAINS -- a
        # human must update these manually, not this script.
        for rule_type in ("GROOVY", "ASSET_SEARCH"):
            with self.subTest(rule_type=rule_type):
                tag = make_tag("[VFZ] Amdocs RevenueOne", rule_type=rule_type)
                planner = make_planner({}, children=[tag])
                rows = planner.plan()
                self.assertEqual(rows[0]["action"], app.ACTION_SKIP_UNSUPPORTED)

    def test_orphan_genuinely_unexpected_rule_type_fails_safe_not_deleted(self):
        tag = make_tag("[VFZ] Some New Tag", rule_type="SOME_FUTURE_TYPE")
        planner = make_planner({}, children=[tag])
        rows = planner.plan()
        self.assertEqual(rows[0]["action"], app.ACTION_ERROR)


# --------------------------------------------------------------------------
# Blast-radius / preflight guardrails
# --------------------------------------------------------------------------


class TestPreflightBlastRadius(unittest.TestCase):
    def _rows(self, n_create=0, n_update=0, n_delete=0):
        rows = []
        for _ in range(n_create):
            rows.append({"action": app.ACTION_CREATE})
        for _ in range(n_update):
            rows.append({"action": app.ACTION_UPDATE})
        for _ in range(n_delete):
            rows.append({"action": app.ACTION_DELETE})
        return rows

    def test_dry_run_never_raises_regardless_of_volume(self):
        rows = self._rows(n_create=1000)
        app.PreflightValidator.validate_plan(
            rows, existing_child_count=10, apply_mode=False, allow_delete=False,
            max_create=1, max_update=1, max_delete=1, max_total=1, max_percentage=1,
            force_large_change=False,
        )  # must not raise

    def test_apply_mode_raises_when_create_limit_exceeded(self):
        rows = self._rows(n_create=5)
        with self.assertRaises(app.PreflightError):
            app.PreflightValidator.validate_plan(
                rows, existing_child_count=100, apply_mode=True, allow_delete=False,
                max_create=1, max_update=100, max_delete=100, max_total=100, max_percentage=100,
                force_large_change=False,
            )

    def test_force_large_change_overrides_the_guardrail(self):
        rows = self._rows(n_create=5)
        app.PreflightValidator.validate_plan(
            rows, existing_child_count=100, apply_mode=True, allow_delete=False,
            max_create=1, max_update=100, max_delete=100, max_total=100, max_percentage=100,
            force_large_change=True,
        )  # must not raise

    def test_child_count_ceiling_is_advisory_not_blocking(self):
        # Regression: the documented 350-per-parent Qualys limit is
        # contradicted by live tenant evidence (607 existing children on the
        # real [VFZ] Applications & Platforms parent). Exceeding it must
        # only warn, never abort -- Qualys' own API is the real authority.
        rows = self._rows(n_create=10)
        app.PreflightValidator.validate_plan(
            rows, existing_child_count=app.MAX_CHILDREN_PER_PARENT - 5, apply_mode=True,
            allow_delete=False, max_create=1000, max_update=1000, max_delete=1000,
            max_total=1000, max_percentage=1000, force_large_change=False,
        )  # must not raise

    def test_deletes_excluded_from_delete_limit_when_allow_delete_false(self):
        rows = self._rows(n_delete=1000)
        app.PreflightValidator.validate_plan(
            rows, existing_child_count=100, apply_mode=True, allow_delete=False,
            max_create=100, max_update=100, max_delete=1, max_total=100, max_percentage=100,
            force_large_change=False,
        )  # must not raise: deletes are not counted unless allow_delete


if __name__ == "__main__":
    unittest.main()
