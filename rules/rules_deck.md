# Layout rules used by `deck_check.json`

- R1 at least a title page and two content pages.
- R2 every page has a title equal to its narrative role plus a verbatim fragment.
- R3 body pages carry at most four bullets; a fifth triggers a recap page.
- R4 every page carries a paragraph id and a verbatim quote that can be re-looked-up.
- R5 quotes are verbatim fragments; no trailing ellipsis.
- R6 speaker notes live in the notes part, never in the slide body.
- R7 role coverage: research question, method, key result, discussion, limitation.
- R8 one PPTX plus one rendered preview per page; previews must match pptx page count.
- R9 figure/table references keep their id; overflow is flagged, never dropped.
- R10 unknown contract keys go to `contract_extras` as null, never invented values.
