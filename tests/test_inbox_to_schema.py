"""
tests/test_inbox_to_schema.py

Regression tests for the inbox → schema pipeline (scripts/inbox_to_schema.py),
focused on adding *new* slots that are assigned to a domain (sub-)class via the
Excel "domain" column.

Background
----------
Previously, a new slot whose "domain" column named a subclass (e.g. anode /
cathode → ElectrochemicalReactor) was silently dropped: it was neither a known
global slot nor present in the class slot_usage, so plan_changes routed it into
the "structural display row" skip branch and _plan_new_slot rejected any
non-empty domain outright. These tests pin the corrected behaviour:

  • a genuinely new domain-assigned slot is planned as a slot_add targeting the
    named subclass and the YAML module that defines it;
  • a real structural slot (own slots:/mixin of the domain class) keeps being
    skipped and is never planned as a new slot;
  • the top-level (empty-domain) path is unchanged;
  • an unrecognised domain is reported as an error rather than written.

The planning-level tests are read-only (plan_changes never writes); the apply
test copies the schema into a tmp dir and redirects SCHEMA_DIR so nothing in the
repository is mutated.
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

import pytest

# ── path setup ──────────────────────────────────────────────────────────────
_ROOT    = Path(__file__).parent.parent
_SCRIPTS = _ROOT / "scripts"
_SCHEMA  = _ROOT / "src" / "coremeta4cat" / "schema"

sys.path.insert(0, str(_SCRIPTS))

import inbox_to_schema as ib                 # noqa: E402
from generate_schema_docs import load_merged_schema  # noqa: E402


# ── helpers ─────────────────────────────────────────────────────────────────

def _slot_row(label: str, domain: str = "", mro: str = "O",
              range_: str = "string", **extra) -> dict:
    """Build one normalised slot row, mirroring parse_excel's output shape."""
    return {
        "label":           label,
        "type":            "slot",
        "domain":          domain,
        "mro":             mro,
        "range":           range_,
        "multivalued":     extra.get("multivalued", ""),
        "inlined_as_list": extra.get("inlined_as_list", ""),
        "unit":            extra.get("unit", ""),
        "uri":             extra.get("uri", ""),
        "description":     extra.get("description", ""),
    }


def _plan(indexes, excel_data):
    """Run the read-only planning phase and return (changes, reporter)."""
    schema, slot_origin, class_origin, label_to_slot, label_to_class = indexes
    reporter = ib.Reporter()
    changes = ib.plan_changes(
        schema, excel_data,
        label_to_slot, label_to_class,
        slot_origin, class_origin,
        reporter,
    )
    return changes, reporter


@pytest.fixture(scope="module")
def indexes():
    """Load the merged schema and build the name/label indexes once."""
    schema = load_merged_schema(str(_SCHEMA))
    slot_origin, class_origin = ib.build_origin_index(_SCHEMA)
    label_to_slot, label_to_class = ib.build_label_index(schema)
    return schema, slot_origin, class_origin, label_to_slot, label_to_class


# ── planning-level tests (read-only) ────────────────────────────────────────

def test_new_slot_with_domain_is_planned_for_subclass(indexes):
    """A new slot assigned to an existing subclass is planned as a slot_add
    targeting that subclass and the YAML module that defines it.

    Uses a synthetic slot name that is never shipped in the schema, so the test
    stays valid even after real domain slots (e.g. anode/cathode) have been
    applied via the pipeline. A hard-coded real name would stop being "new" once
    it lands in the schema and would silently invalidate the assertion.
    """
    schema        = indexes[0]
    label_to_slot = indexes[3]

    label = "qa synthetic electrode probe"
    name  = "qa_synthetic_electrode_probe"

    # Precondition: the slot must genuinely be absent for this test to be
    # meaningful. Asserting it makes the assumption explicit and self-checking.
    assert name not in schema.get("slots", {})
    assert label not in label_to_slot
    assert name not in ib.get_all_class_slots(schema, "ElectrochemicalReactor")

    excel = {"Reaction": [
        _slot_row(label, domain="ElectrochemicalReactor", mro="M"),
    ]}
    changes, reporter = _plan(indexes, excel)

    adds = [c for c in changes if c["type"] == "slot_add" and c["name"] == name]
    assert len(adds) == 1, f"expected exactly one slot_add for '{name}'"
    add = adds[0]
    assert add["schema_class"] == "ElectrochemicalReactor"
    assert add["mro"] == "M"
    assert Path(add["_target"]).name == "coremeta4cat_reaction_ap.yaml"
    assert not reporter.has_errors


def test_structural_subclass_slot_is_not_added(indexes):
    """A slot that already belongs to the domain class -- via the class's own
    slots: list or via a mixin -- keeps being skipped and is never planned as a
    new slot."""
    # 'title'/'description' are in QuantitativeRange.slots:; 'has concentration'
    # reaches CoPrecipitation via PrecipitationMixin. All three are real
    # structural rows emitted by the Excel generator and must stay skips.
    excel = {
        "Reaction": [
            _slot_row("title", domain="QuantitativeRange"),
            _slot_row("description", domain="QuantitativeRange"),
        ],
        "Synthesis": [
            _slot_row("has concentration", domain="CoPrecipitation"),
        ],
    }
    changes, reporter = _plan(indexes, excel)

    new_names = {c["name"] for c in changes if c["type"] == "slot_add"}
    assert "title" not in new_names
    assert "description" not in new_names
    assert "has_concentration" not in new_names
    assert not reporter.has_errors


def test_top_level_new_slot_still_targets_sheet_class(indexes):
    """Regression guard: an empty-domain new slot still attaches to the sheet's
    top-level data class (CatalyticReaction for the Reaction sheet)."""
    excel = {"Reaction": [
        _slot_row("brand new toplevel field", domain="", mro="R"),
    ]}
    changes, reporter = _plan(indexes, excel)

    add = next(
        (c for c in changes
         if c["type"] == "slot_add" and c["name"] == "brand_new_toplevel_field"),
        None,
    )
    assert add is not None
    assert add["schema_class"] == "CatalyticReaction"
    assert not reporter.has_errors


def test_new_slot_with_unknown_domain_is_an_error(indexes):
    """A new slot whose domain is not a known class is reported as an error and
    not planned for application."""
    excel = {"Reaction": [
        _slot_row("weird slot", domain="NotARealClass"),
    ]}
    changes, reporter = _plan(indexes, excel)

    assert reporter.has_errors
    assert not any(
        c["type"] == "slot_add" and c["name"] == "weird_slot" for c in changes
    )


# ── apply-level test (writes into a tmp copy, never the repo) ────────────────

def test_apply_writes_new_domain_slot_into_subclass(tmp_path, monkeypatch):
    """End-to-end: applying a domain-assigned slot_add creates the global slot
    definition and references it from the subclass's slots: list."""
    dst = tmp_path / "schema"
    dst.mkdir()
    for f in _SCHEMA.glob("*.yaml"):
        shutil.copy(f, dst / f.name)
    monkeypatch.setattr(ib, "SCHEMA_DIR", dst)

    reporter = ib.Reporter()
    schema = load_merged_schema(str(dst))
    slot_origin, class_origin = ib.build_origin_index(dst)
    label_to_slot, label_to_class = ib.build_label_index(schema)

    excel = {"Reaction": [
        _slot_row("test electrode", domain="ElectrochemicalReactor", mro="M",
                  description="A test electrode slot."),
    ]}
    changes = ib.plan_changes(
        schema, excel, label_to_slot, label_to_class,
        slot_origin, class_origin, reporter,
    )

    # Apply only our slot_add so unrelated deletion-detection changes (the
    # single-row workbook makes every other Reaction slot look "missing") do
    # not mutate the copy.
    adds = [c for c in changes
            if c["type"] == "slot_add" and c["name"] == "test_electrode"]
    assert len(adds) == 1
    ib.apply_changes(adds, reporter)

    doc = ib._load_yaml(dst / "coremeta4cat_reaction_ap.yaml")

    # (a) global slot definition was created with the required flag
    assert "test_electrode" in (doc.get("slots") or {})
    assert doc["slots"]["test_electrode"].get("required") is True

    # (b) the subclass references the new slot in its slots: list
    er = (doc.get("classes") or {}).get("ElectrochemicalReactor") or {}
    assert "test_electrode" in (er.get("slots") or [])
