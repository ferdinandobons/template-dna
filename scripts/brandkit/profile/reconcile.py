# SPDX-License-Identifier: MIT
"""Kind-agnostic reconciliation policy shared by the generators (plan §6).

The destructive-action floor lives here, in the profile layer, because it is a
single brand-wide policy: a DESTRUCTIVE verdict (a ``clear`` on a slot
determinism does not also corroborate) is downgraded to KEEP + WARNING unless
the model's ``comprehension.confidence`` clears the floor. Additive FILL is never
gated on confidence - a wrong fill is recoverable, a wrong delete is not. The
docx and pptx cover reconcilers both route their confidence check through this
one definition so the threshold and comparison sense live in **one** place.
"""

from __future__ import annotations

# Below this confidence the model's DESTRUCTIVE verdicts are downgraded to
# KEEP + WARNING (the destructive-action floor, plan §6).
DESTRUCTIVE_CONFIDENCE_FLOOR: float = 0.5


def confidence_clears_floor(confidence: float) -> bool:
    """Whether ``confidence`` is high enough to honor a DESTRUCTIVE verdict."""
    return confidence >= DESTRUCTIVE_CONFIDENCE_FLOOR
