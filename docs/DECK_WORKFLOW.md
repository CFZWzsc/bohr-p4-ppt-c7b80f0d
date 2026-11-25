# Deck workflow (single run)

1. `pre_apply`: if the input directory contains a `.pptx`/`.potx` template, read its
   slide size, theme minor font and colour scheme, and inherit the slide size and
   theme fonts into the OOXML writer. What is not inherited (colour scheme, master
   background shapes) is recorded in `template_probe.json` instead of claimed.
2. content selection: every page gets one claim taken verbatim from the material,
   plus at most four verbatim bullets, each carrying its paragraph id.
3. figure registry: figure/table references in the material become structured slots
   with ids; a third slot in one section is marked `placement=overflow`.
4. build: `deck.pptx` (real OOXML) + one SVG/PNG preview per page + speaker notes in
   `notesSlide` only.
5. self-check: `deck_check.json` (D01-D10) and `metrics.json` record what passed.
6. gain record: `baseline_comparison.md` compares against a reproducible naive
   baseline on the same input.
