# Error handling

| situation | behaviour |
|---|---|
| no `case.json` | honest `input_error` result, contract keys present, exit 0 |
| template missing | built-in theme; `template_probe.json` says `template_provided=false` |
| template unreadable | `template_probe.json` records the parse failure; built-in theme used |
| material file unreadable | listed in `coverage.not_read`; page count never padded |
| figure source missing | slot kept with its caption/id, marked as unavailable |
| text box overflow | bullets are clipped verbatim (never paraphrased) and the page is split |
| export renderer missing | SVG vector sources kept; PNG absence recorded, never faked |
| any unhandled exception | traceback to `work/entry-error.log`, partial result written, exit 0 |
