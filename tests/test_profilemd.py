# SPDX-License-Identifier: MIT
"""Shared PROFILE.md sections (common/profilemd) - the authoring surface.

These guards pin the format-uniform authoring sections: the role table, the
semantic palette-role table (tokens an author may name), and the authoring
hints (role-first, fragments-aware). All three extract writers render through
this one module, so the three formats' PROFILE.md can never drift apart.
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from brandkit.common import profilemd  # noqa: E402


def _profile(**over) -> dict:
    base = {
        "identity": {"display_name": "t"},
        "roles": {
            "_index": ["paragraph", "heading.1"],
            "paragraph": {
                "resolver": {"style_name": "Normal"},
                "status": "robust",
                "usage": {"scope": "body", "placement": "freeform"},
            },
            "heading.1": {
                "resolver": {"style_name": "Heading 1"},
                "status": "robust",
                "usage": {
                    "scope": "body",
                    "placement": "structural",
                    "required": True,
                    "order": 2,
                },
            },
        },
        "theme": {
            "colors": {"accent1": {"hex": "2B7CD3"}, "dk1": {"hex": "16213F"}},
            "palette_roles": {"primary": {"theme": "accent1"}},
        },
        "comprehension": {"status": "absent", "fragments": []},
    }
    base.update(over)
    return base


class ProfileMdTest(unittest.TestCase):
    def test_roles_md_lists_every_indexed_role_with_usage(self):
        text = "\n".join(profilemd.roles_md(_profile()))
        self.assertIn("## Roles", text)
        self.assertIn("`paragraph`: Normal (robust)", text)
        self.assertIn("`heading.1`: Heading 1 (robust)", text)
        self.assertIn("required · order=2", text)

    def test_roles_md_empty_index_renders_nothing(self):
        self.assertEqual(profilemd.roles_md(_profile(roles={"_index": []})), [])

    def test_palette_roles_md_resolves_hex_and_forbids_raw(self):
        text = "\n".join(profilemd.palette_roles_md(_profile()))
        self.assertIn("## Brand palette roles", text)
        self.assertIn("`primary` -> `accent1` (#2B7CD3)", text)
        self.assertIn("never a raw hex", text)

    def test_palette_roles_md_absent_renders_nothing(self):
        prof = _profile(theme={"colors": {}, "palette_roles": {}})
        self.assertEqual(profilemd.palette_roles_md(prof), [])

    def test_authoring_hints_advertise_fragments_when_present(self):
        prof = _profile(
            comprehension={
                "status": "present",
                "fragments": [{"name": "kpi.card"}, {"name": "two.col"}],
            }
        )
        text = "\n".join(profilemd.authoring_hints_md(prof))
        self.assertIn("REUSE the comprehended fragments (kpi.card, two.col)", text)
        self.assertIn("{{slot}}", text)

    def test_authoring_hints_point_to_comprehend_when_absent(self):
        text = "\n".join(profilemd.authoring_hints_md(_profile()))
        self.assertIn("## Authoring hints", text)
        self.assertIn("After `comprehend`", text)
        self.assertIn("Never name a style, font,", text)


if __name__ == "__main__":
    unittest.main()
