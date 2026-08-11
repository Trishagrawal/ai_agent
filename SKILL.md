---
name: afg-ppt-builder
description: Build and amend AFG-branded PowerPoint decks from AFG_template.pptx. Use whenever the user asks for a presentation, deck, or slides, or asks to change/add to a deck already built in this conversation. Never write ad hoc python-pptx code for this — always call the bundled script.
---

# AFG PPT Builder

This skill produces AFG-branded `.pptx` files by calling a **fixed,
deterministic script** (`scripts/afg_deck_builder.py`) via Code Interpreter.
Do not write your own placeholder-filling logic, and do not generate raw
slide XML by hand — every past attempt at that produced broken output
(content dumped into the wrong placeholder, duplicated content across
slides, text written into picture placeholders, blank filler images). The
script exists specifically to make those mistakes structurally impossible.

**Read this before planning any slide: `AFG_template.pptx` is a 176-slide
reference library, not just the ~34 base layouts (Title Page, Content
Slide_N columns, etc.).** Those 34 are the structural skeleton every
slide needs — they are NOT the full set of design options. Sections like
"Diagrams and Charts," "Next Steps," "Numbers and factsheets," and
"Organigram Slides" contain hand-designed illustrations built from real
shapes, meant to be cloned directly rather than approximated. **The
complete slide-by-slide index of all 176 is inline further down this
file** ("All 176 template slides" under "Layout and placeholder
reference") — check it before defaulting to a plain layout or a
procedural diagram, every time, not only when something already looks
repetitive. This has been the single most repeated correction on this
skill — a deck that only ever uses the 34 base layouts, ignoring the
other 142 reference slides, is treated as a planning miss, not an
acceptable default.

## The template is bundled with this skill — no fetching needed

`AFG_template.pptx` ships inside this skill package itself, in the same
directory as this `SKILL.md`. Copilot Studio's Agent Skills feature loads
bundled resource files as part of the skill — when this skill is active,
the template is directly available in the working files, the same way
`scripts/afg_deck_builder.py` is. This is the reliable, automatic path:

- **Do not** ask the user to attach the template to the chat.
- **Do not** rely on a general knowledge-source entry for the template —
  knowledge sources get chunked/indexed for text retrieval, not exposed
  as raw file bytes to Code Interpreter; this bundled copy is what
  actually gets used.
- Reference it by its path within the skill package (e.g.
  `AFG_template.pptx`, resolved relative to this skill's own directory)
  when calling `afg_deck_builder.py build --template ...`.
- If, in your specific Copilot Studio setup, the bundled file isn't
  showing up in Code Interpreter's working files for some reason, that's
  a setup issue worth checking directly (re-upload the skill package,
  confirm the skill is actually active/selected for this request) rather
  than something to work around by asking the user for the file — the
  file is supposed to always be there once the skill loads correctly.
- Keeping the template inside the skill package also means updates are
  simple: if the AFG brand template changes, re-upload a new version of
  this same skill zip with the updated `.pptx` inside it, and every
  future generation automatically uses the new version — no separate
  file-management step anywhere else.

### The template's reference library — the complete map of all 176 slides

`AFG_template.pptx` (as authored) contains 176 example slides organized
into 21 real PowerPoint Sections (visible in PowerPoint's Slide Sorter
view) — this is the same underlying template file, so this map is exact,
not approximate.

**All 176 ARE directly editable — every one of them, not just the 34
most commonly used.** Two distinct mechanisms make this true, and it
matters to know which applies to what you're doing:
- **Structurally, all 176 are built on the same 34 real PowerPoint
  layouts** (Section "Layout and placeholder reference" below) — this is
  a verifiable fact about the file (`python-pptx` reports exactly 34
  entries in `slide_layouts`), not a limitation this skill imposes. A
  reference slide titled "Circle Segments" and a plain content slide can
  both be built on, say, `Content Slide_empty_light` — the illustration
  is extra hand-designed shapes added on top of that shared layout, not
  a different layout underneath it.
- **Functionally, `clone_slide` (see "Cloned visuals" below) lets you
  take ANY of the 176 directly and edit its actual content in place** —
  its own title/subheadline through normal placeholders, and every one
  of its decorative shapes' text through `text_replacements`. This is
  full editing of the real slide, not an approximation of it and not
  limited to extracting one piece — verified working end-to-end on
  slides 105 and 120 (see "Verified examples"), covering both a
  multi-piece illustration (105's ring plus four separate corner
  callouts) and a connected-chain illustration (120's five linked
  badges). Nothing in the mechanism is specific to those two slides —
  any of the 176 can be targeted the same way; 105/106/120 are simply
  the ones already confirmed clean by an actual render, not the only
  ones the tool supports.

`_clear_existing_slides` strips every one of these 176 reference slides
(and the section grouping itself) from the in-memory copy before any new
content is added — the bundled template file on disk is never modified,
so this reference library is intact again the next time the skill loads,
and none of the 176 is ever copied wholesale into a generated deck
without going through `clone_slide`/`cloned_visual` and having its
placeholder text replaced.

| Section (exact slide range) | What's there | How to use it |
|---|---|---|
| Masterslides (1-32) | Raw layout/master definitions | Template-authoring artifact, not a content reference — ignore. |
| Corporate Introduction layouts (33-45) | Company/team intro slides: story, leadership, facts+icons | Reference for "About us"/intro sections — content maps to existing layouts + `card_grid`/`accent_list`. |
| Executive Summary and resourcing (46-48) | Summary + resourcing slides | `card_grid` or `stat_row` for the summary; plain content for resourcing detail. |
| Strategy & Outcome (49-51) | Strategy/outcome statements | `process_steps` or `card_grid` depending on whether it's sequential or parallel. |
| Quarterly Roadmap (52-53) | Quarter-by-quarter plan | `process_steps`, or `cloned_visual` of slide 105/106 if it's proportional rather than sequential. |
| Brands (54-57) | Logo/brand asset slides | Template-authoring artifact — ignore unless the user specifically wants a logo wall. |
| Text Content (58-73) | Plain text/paragraph slide variants | Covered by existing `Content Slide_N columns` + plain `"bullets"`. |
| Text and images (74-86) | Text paired with images, various ratios | Covered by `Split Content image` — pick the variant by dark/light, not by these sub-styles. |
| Table Slides (87-94) | Tables and timetables | No native table support yet — fall back to `card_grid` (one card per row) or tell the user tables aren't supported, rather than improvising. |
| Organigram Slides (95-99) | Org charts / hierarchy | `card_grid` (one card per role/box) or nested `process_steps` for a reporting chain. |
| Numbers and factsheets (100-103) | KPI/stat callouts | `stat_row` for headline numbers; `card_grid` if each number needs its own descriptive card. |
| Diagrams and Charts (104-114) | Circle segments, generic diagrams, bar charts, keynumbers | **Slides 105/106 (Circle Segments) are verified cloneable — see "Cloned visuals" below.** Slide 114 is a native chart — model your own chart on it, don't clone a chart object. |
| Next Steps (115-120) | Numbered step chains, various layouts | **Slide 120 verified via `clone_slide` (whole-slide mode) — see "Cloned visuals" below. Single-shape `cloned_visual` mode is broken for it (rescaling distorts the badges); slides 116-119 aren't single groups and aren't individually verified either way — catalog+render before trusting them, or use `process_steps` instead.** |
| Division Style/Theme (121) | Single divider slide | Style reference only. |
| Automotive Style (122-127) | Title/Chapter/Split Content slides re-skinned for Automotive | Not unique illustrations — these are the same 34 layouts. Signal to pair with the `automotive-*` icon set instead of the generic one. |
| Retail Style (128-134) | Same pattern, Retail | Pair with `retail-*` icons. |
| Real Estate Style (135-141) | Same pattern, Real Estate | Pair with `real-estate-*` icons. |
| Finance Style (142-148) | Same pattern, Finance | Pair with `financial-services-*` icons. |
| Health Style (149-155) | Same pattern, Health | No dedicated icon set yet — use the generic `icons/` set. |
| Education Style (156-162) | Same pattern, Education | No dedicated icon set yet — use the generic `icons/` set. |
| Corporate Style (163-169) | Same pattern, Corporate | Pair with `corporate-*` icons. |
| Blue Style (170-176) | Same pattern, generic blue | Use the generic `icons/` set. |

**In practice: sections 33-120 are where a genuinely useful illustration
to clone might live (Diagrams and Charts, Next Steps especially); 121-176
are a style/icon-pairing signal, not new structural content — don't spend
time cataloging individual slides in that range.** If a request doesn't
map to anything distinctive above, it's normal content covered by the 34
layouts directly — this table exists to catch the cases where a
hand-designed illustration would look better than this skill's own
procedural approximation of the same idea, not to imply every one of the
176 slides needs individual consideration every time.


## Your job vs. the script's job

- **You (the agent):** gather the user's story/content, ask clarifying
  questions, propose a slide-by-slide outline, get explicit approval, then
  translate the approved outline into the JSON plan format below.
- **The script:** opens the template, fills each placeholder by its exact
  idx, checks each placeholder's type before writing to it, deletes any
  placeholder you didn't address, checks for accidental duplicate content
  across slides, and saves the file. It also handles amending an existing
  file in place.

Always call the script through Code Interpreter. Never hand-write the
fill/cleanup logic yourself, even for a "quick" one-slide change.

## Session state — always amend, never rebuild from scratch

The first time you build a deck in a conversation, remember the exact
output path you saved it to. For every later change request in the same
conversation ("change slide 4", "add a slide about X", "swap that image"),
call the script's `amend` command against that same file — never call
`build` again from the original brief. `build` recreates the whole file
from nothing; `amend` edits only what you target and leaves everything
else byte-for-byte the same. Re-generating with `build` on every edit
silently reverts changes the user already approved.

### Matching an existing deck's style when adding slides — new content must look native, not bolted-on

When a request adds slides to a deck that already exists (a bigger ask —
"add 15 more slides," "expand this into a full deck," not a small
one-slide tweak), the new slides must match the visual and content
treatment the existing slides already established — never noticeably
downgrade to something plainer just because it's faster to generate. This
happened in practice: an expansion from 11 to 30 slides used only plain
bulleted text and a generic stock background image for all 19 new
slides, while the original 11 slides used `accent_list` icons, a cloned
Circle Segments diagram, and 4-column icon grids — the new slides read as
a visibly different, lower-effort deck grafted onto the front of a
better one, not a continuation of it.

**Before writing new slides for an expansion, look at what the existing
similar-content slides actually did**, then match that level of
treatment:
- If existing "use case" slides used `accent_list` with icons or
  `card_grid`, new use-case slides should too — not fall back to plain
  `"bullets"`.
- If existing slides used a specific `bullet_style`, don't silently
  revert to the default `"dot"` for new ones.
- If the deck already established a rhythm of diagrams/charts appearing
  regularly, new slides should continue that rhythm, not go bullet-only
  for a long stretch.

**Never copy a slide's own bullet/description text as the starting
point for a new slide, even about a related topic.** This is exactly how
a much worse version of the "boring, mismatched" problem happened: the
identical 3-bullet block ("Reduce manual effort" / "Improve visibility"
/ "Enable mobile execution") got reused verbatim across 9 different
topically-distinct new slides — a generic filler pool substituted for
real per-slide writing. The build script's `check_duplicate_content`
hard-blocks this exact pattern for body content now (see below) — but
the fix is writing genuinely per-slide content in the first place, not
relying on the block to catch it after the fact.

## Layout and placeholder reference

### Quick decision guide — pick by what the content actually is, not by habit

Work through this list for every slide before defaulting to whatever
layout the last few slides used. This is the single fastest fix for a
deck that technically varies layouts but still feels repetitive — most
repetition comes from never checking this list, not from these options
being unavailable.

| The content is... | Reach for... |
|---|---|
| A phased plan or proportional/parallel breakdown (rollout stages, capability pillars, budget split) | **`clone_slide`** of slide 105 or 106 (verified — see "Cloned visuals" below) |
| A sequential chain of numbered steps that needs a hand-designed look, not just `process_steps` | **`clone_slide`** of slide 120 (verified) |
| A sequence with a clear order (workflow, "how it works", roadmap) | **`process_steps`** diagram |
| Parallel categories, each with its own short list | **`card_grid`** diagram (one call per column if the layout already splits into columns) |
| A handful of feature/capability highlights, optionally each with a relevant icon | **`accent_list`** diagram |
| Illustrative headline metrics ("70% faster") without a full data set | **`stat_row`** diagram |
| Real numeric data — a trend, a comparison, a breakdown | a native **chart** (line/bar/pie — match type to the data, Charts section below), not a diagram |
| A stakeholder or customer quote | **Quote Slide 1 or 2** |
| A natural break between major sections of the deck | **Chapter Slide** (only if real content follows it) |
| An explicit two-way comparison or before/after | **Split design** |
| One strong single visual moment (a launch, a location, a hero shot) | **Full image Slide** |
| Text paired with one supporting photo/illustration | **Split Content image** |
| A short statement with no bullets needed | **Content Slide_empty** |
| Prose/bullets that doesn't reduce to any of the above | plain `"bullets"` in **Content Slide_1/2/3/4 columns** (still vary the column count/theme rather than defaulting to the same one every time) |

None of these are exotic edge cases — they're the normal, expected range
of a deck built from this template. A deck that only ever reaches for
the last row is the failure mode this table exists to prevent.

### All 176 template slides — the complete list, right here in this file

**This table is loaded automatically as part of this skill's core
instructions — it is not a separate file you need to decide to open.**
Check it against every slide before defaulting to a plain layout,
a procedural diagram, or bullets written from scratch. Same
three-column format as the 34-layout table further down this section.
**Priority order, explicitly: check the row below FIRST.** If a row
matches the slide's actual content, clone it (`"clone_slide": <#>` — see
"Cloned visuals" below for the mechanics: it edits that slide's real
content directly, it does not just reference or copy an illustration out
of it). Only fall back to the shorter 34-layout table beneath this one,
a procedural diagram, or bullets from scratch when nothing here is a
genuine match for the content — that fallback path is not the default,
this table is. (The same content is also duplicated standalone in
`reference_library_index.md`, bundled alongside this file, purely for
quick command-line lookup during a session — this copy, inline here, is
the one that's guaranteed to already be part of what you know before
you start planning.)

**Real photography inside a cloned reference slide stays exactly as
designed.** `clone_slide` clones every shape on the slide, including any
actual photo already embedded in it — leave it as part of the
composition and only supply `text_replacements` for the text. Don't
strip an embedded photo out, and don't treat it as something that needs
replacing — only swap it if the user explicitly asks for a different
image on that specific slide.

**Every single one of these 176 is technically clonable — there is no
code-level restriction on any slide number, including Masterslides
(1-32), Brands (54-57), or the division Style sections (121-176).**
Where a row says "template-authoring artifact" or "use the standard
layout instead," that's a recommendation about what looks best, not a
barrier the script enforces — cloning any of those slides works exactly
like any other row if you have a real reason to. Don't read those notes
as "this range will fail" and avoid the reference table more broadly
because of it.

| Slide (clone_slide: #) | Use for | Real placeholder structure |
|---|---|---|
| **Slide 1** — Masterslides — Simple text layout | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subtitle(idx1); 1 standalone text box(es); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 2** — Masterslides — Simple text layout [#2] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subtitle(idx1); 1 standalone text box(es); built on layout `Title Page_dark` |
| **Slide 3** — Masterslides — Simple text layout [#3] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subtitle(idx1); 1 standalone text box(es); built on layout `Title Page_light` |
| **Slide 4** — Masterslides — Simple text layout [#4] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subtitle(idx1); built on layout `Title Page_dark` |
| **Slide 5** — Masterslides — Simple text layout [#5] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subtitle(idx1); built on layout `Title Page_light` |
| **Slide 6** — Masterslides — 4-item text layout | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_dark` |
| **Slide 7** — Masterslides — 4-item text layout [#7] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_light` |
| **Slide 8** — Masterslides — 3-item text layout | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subtitle(idx1); 3 standalone text box(es); built on layout `Title Page with image_dark` |
| **Slide 9** — Masterslides — 3-item text layout [#9] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subtitle(idx1); 3 standalone text box(es); built on layout `Title Page with image_light` |
| **Slide 10** — Section separator slides (simple text layout) | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subtitle(idx1); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 11** — Masterslides — Simple text layout [#11] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_dark` |
| **Slide 12** — Masterslides — Simple text layout [#12] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_light` |
| **Slide 13** — Blank slides (simple text layout) | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subtitle(idx1); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 14** — Masterslides — Simple text layout [#14] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subheadline(idx13); built on layout `Content Slide_empty_dark` |
| **Slide 15** — Masterslides — Simple text layout [#15] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 16** — Masterslides — Simple text layout [#16] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subheadline(idx13); built on layout `Split Content image_dark_1` |
| **Slide 17** — Masterslides — Simple text layout [#17] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_2` |
| **Slide 18** — Masterslides — Simple text layout [#18] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_3` |
| **Slide 19** — Masterslides — Simple text layout [#19] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split design` |
| **Slide 20** — Full page image options (simple text layout) | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subtitle(idx1); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 21** — Masterslides — Simple text layout [#21] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | built on layout `Full image Slide_dark_1` |
| **Slide 22** — Masterslides — Simple text layout [#22] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | built on layout `Full image Slide_dark_2` |
| **Slide 23** — Masterslides — Simple text layout [#23] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | built on layout `Full image Slide_light` |
| **Slide 24** — Quote slides (simple text layout) | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subtitle(idx1); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 25** — Masterslides — Simple text layout [#25] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | attribution(idx13) + quote text(idx14); built on layout `Quote Slide 1_dark` |
| **Slide 26** — Masterslides — Simple text layout [#26] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | attribution(idx13) + quote text(idx14); built on layout `Quote Slide 1_light` |
| **Slide 27** — Masterslides — Simple text layout [#27] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | attribution(idx13) + quote text(idx14); built on layout `Quote Slide 2_dark` |
| **Slide 28** — Masterslides — Simple text layout [#28] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | attribution(idx13) + quote text(idx14); built on layout `Quote Slide 2_light` |
| **Slide 29** — Thank you slides (simple text layout) | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subtitle(idx1); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 30** — Masterslides — 1-group diagram (0 text elements) | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subheadline(idx13); 1 standalone text box(es); 1 illustration group(s) (group of 0 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Thank you and contact_dark` |
| **Slide 31** — Masterslides — Simple text layout [#31] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subheadline(idx13); built on layout `Thank you and contact_dark` |
| **Slide 32** — Masterslides — Simple text layout [#32] | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subheadline(idx13); built on layout `Thank you and contact_light` |
| **Slide 33** — Corporate Introduction layouts (simple text layout) | Simple content slide under 'Corporate Introduction' — clone via `clone_slide: 33`, or use the standard layout below directly for the same placeholder structure without the extra decoration. | title(idx0); subtitle(idx1); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 34** — Our Story (17-item icon grid (14 icons, caption per icon)) | Data/illustration slide — 17-item icon grid (14 icons, caption per icon), 57 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 34` to reuse this exact hand-designed composition for matching content in 'Corporate Introduction'. | title(idx0); subheadline(idx13); 43 standalone text box(es); 17 illustration group(s) (group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s)); 14 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_dark` |
| **Slide 35** — Leadership slide (16-item text/number grid) | Data/illustration slide — 16-item text/number grid, 31 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 35` to reuse this exact hand-designed composition for matching content in 'Corporate Introduction'. | title(idx0); subheadline(idx13); 16 standalone text box(es); 15 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_dark` |
| **Slide 36** — Corporate Introduction — 4-group diagram (3 text elements) | Data/illustration slide — 4-group diagram (3 text elements), 4 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 36` to reuse this exact hand-designed composition for matching content in 'Corporate Introduction'. | 1 standalone text box(es); 4 illustration group(s) (group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 0 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Full image Slide_dark_2` |
| **Slide 37** — Text and image (5-item text layout) | Data/illustration slide — 5-item text layout, 5 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 37` to reuse this exact hand-designed composition for matching content in 'Corporate Introduction'. | title(idx0); subheadline(idx13); 5 standalone text box(es); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_light_2` |
| **Slide 38** — Corporate Introduction — 1-item icon grid (9 icons, caption per icon) | Data/illustration slide — 1-item icon grid (9 icons, caption per icon), 19 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 38` to reuse this exact hand-designed composition for matching content in 'Corporate Introduction'. | title(idx0); subheadline(idx13); 10 standalone text box(es); 1 illustration group(s) (group of 0 text(s)); 9 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_dark` |
| **Slide 39** — Corporate Introduction — 1-item icon grid (9 icons, caption per icon) [#39] | Data/illustration slide — 1-item icon grid (9 icons, caption per icon), 19 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 39` to reuse this exact hand-designed composition for matching content in 'Corporate Introduction'. | title(idx0); subheadline(idx13); 10 standalone text box(es); 1 illustration group(s) (group of 0 text(s)); 9 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_dark` |
| **Slide 40** — Facts and icons (4-group diagram (0 text elements)) | Data/illustration slide — 4-group diagram (0 text elements), 11 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 40` to reuse this exact hand-designed composition for matching content in 'Corporate Introduction'. | title(idx0); subheadline(idx13); 11 standalone text box(es); 4 illustration group(s) (group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_dark` |
| **Slide 41** — Facts and icons (5-item callout row with 4 icon(s)) | Data/illustration slide — 5-item callout row with 4 icon(s), 9 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 41` to reuse this exact hand-designed composition for matching content in 'Corporate Introduction'. | title(idx0); subheadline(idx13); 5 standalone text box(es); 4 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 42** — Corporate Introduction — 1-group diagram (0 text elements) | Data/illustration slide — 1-group diagram (0 text elements), 6 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 42` to reuse this exact hand-designed composition for matching content in 'Corporate Introduction'. | title(idx0); subheadline(idx13); 6 standalone text box(es); 1 illustration group(s) (group of 0 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 43** — Corporate Introduction — 9-item text/number grid | Data/illustration slide — 9-item text/number grid, 9 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 43` to reuse this exact hand-designed composition for matching content in 'Corporate Introduction'. | title(idx0); subheadline(idx13); 9 standalone text box(es); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 44** — Corporate Introduction — 1-item icon grid (3 icons, caption per icon) | Data/illustration slide — 1-item icon grid (3 icons, caption per icon), 12 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 44` to reuse this exact hand-designed composition for matching content in 'Corporate Introduction'. | title(idx0); subheadline(idx13); 9 standalone text box(es); 1 illustration group(s) (group of 0 text(s)); 3 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 45** — Corporate Introduction — 12-item text/number grid | Data/illustration slide — 12-item text/number grid, 16 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 45` to reuse this exact hand-designed composition for matching content in 'Corporate Introduction'. | title(idx0); subheadline(idx13); 12 standalone text box(es); 4 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 46** — Executive Summary  and resourcing (simple text layout) | Simple content slide under 'Executive Summary' — clone via `clone_slide: 46`, or use the standard layout below directly for the same placeholder structure without the extra decoration. | title(idx0); subtitle(idx1); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 47** — Executive Summary (3-item icon grid (4 icons, caption per icon)) | Data/illustration slide — 3-item icon grid (4 icons, caption per icon), 16 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 47` to reuse this exact hand-designed composition for matching content in 'Executive Summary'. | title(idx0); subheadline(idx13); 12 standalone text box(es); 3 illustration group(s) (group of 0 text(s), group of 0 text(s), group of 0 text(s)); 4 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 48** — Resourcing (text + 1 image layout) | Data/illustration slide — text + 1 image layout, 1 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 48` to reuse this exact hand-designed composition for matching content in 'Executive Summary'. | title(idx0); subheadline(idx13); 1 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 49** — Strategy & Outcome (simple text layout) | Simple content slide under 'Strategy & Outcome' — clone via `clone_slide: 49`, or use the standard layout below directly for the same placeholder structure without the extra decoration. | title(idx0); subtitle(idx1); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 50** — Strategy & Outcome (5-group diagram (0 text elements)) | Data/illustration slide — 5-group diagram (0 text elements), 10 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 50` to reuse this exact hand-designed composition for matching content in 'Strategy & Outcome'. | title(idx0); subheadline(idx13); 10 standalone text box(es); 5 illustration group(s) (group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 51** — Strategy & Outcome (5-group diagram (0 text elements)) [#51] | Data/illustration slide — 5-group diagram (0 text elements), 10 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 51` to reuse this exact hand-designed composition for matching content in 'Strategy & Outcome'. | title(idx0); subheadline(idx13); 10 standalone text box(es); 5 illustration group(s) (group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 52** — Quarterly Roadmap (simple text layout) | Simple content slide under 'Quarterly Roadmap' — clone via `clone_slide: 52`, or use the standard layout below directly for the same placeholder structure without the extra decoration. | title(idx0); subtitle(idx1); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 53** — Quarterly Roadmap (simple text layout) [#53] | Simple content slide under 'Quarterly Roadmap' — clone via `clone_slide: 53`, or use the standard layout below directly for the same placeholder structure without the extra decoration. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 54** — Brands and Logos (simple text layout) | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subtitle(idx1); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 55** — Logos (simple text layout) | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 56** — Logo (connected multi-part illustration (12 text elements across 2 group(s))) | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subheadline(idx13); 2 illustration group(s) (group of 9 text(s), group of 3 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 57** — Logo (3-group diagram (15 text elements)) | Template-authoring artifact (raw layout/master or logo asset) — not content to build a slide from. | title(idx0); subheadline(idx13); 10 standalone text box(es); 3 illustration group(s) (group of 5 text(s), group of 5 text(s), group of 5 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 58** — Text Content (simple text layout) | Simple content slide under 'Text Content' — clone via `clone_slide: 58`, or use the standard layout below directly for the same placeholder structure without the extra decoration. | title(idx0); subtitle(idx1); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 59** — Text Content — 9-item text/number grid | Data/illustration slide — 9-item text/number grid, 9 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 59` to reuse this exact hand-designed composition for matching content in 'Text Content'. | title(idx0); subheadline(idx13); 9 standalone text box(es); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 60** — Text Content — 3-group diagram (0 text elements) | Data/illustration slide — 3-group diagram (0 text elements), 8 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 60` to reuse this exact hand-designed composition for matching content in 'Text Content'. | title(idx0); subheadline(idx13); 8 standalone text box(es); 3 illustration group(s) (group of 0 text(s), group of 0 text(s), group of 0 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 61** — Text Content — 2-group diagram (8 text elements) | Data/illustration slide — 2-group diagram (8 text elements), 9 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 61` to reuse this exact hand-designed composition for matching content in 'Text Content'. | title(idx0); subheadline(idx13); 1 standalone text box(es); 2 illustration group(s) (group of 4 text(s), group of 4 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 62** — Text Content — 4-group diagram (12 text elements) | Data/illustration slide — 4-group diagram (12 text elements), 16 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 62` to reuse this exact hand-designed composition for matching content in 'Text Content'. | title(idx0); subheadline(idx13); 4 standalone text box(es); 4 illustration group(s) (group of 3 text(s), group of 3 text(s), group of 3 text(s), group of 3 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 63** — Text Content — 4-group diagram (12 text elements) [#63] | Data/illustration slide — 4-group diagram (12 text elements), 16 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 63` to reuse this exact hand-designed composition for matching content in 'Text Content'. | title(idx0); subheadline(idx13); 4 standalone text box(es); 4 illustration group(s) (group of 3 text(s), group of 3 text(s), group of 3 text(s), group of 3 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 64** — Text Content — 16-item text/number grid | Data/illustration slide — 16-item text/number grid, 16 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 64` to reuse this exact hand-designed composition for matching content in 'Text Content'. | title(idx0); subheadline(idx13); 16 standalone text box(es); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 65** — Text Content — 4-group diagram (0 text elements) | Data/illustration slide — 4-group diagram (0 text elements), 7 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 65` to reuse this exact hand-designed composition for matching content in 'Text Content'. | title(idx0); subheadline(idx13); 7 standalone text box(es); 4 illustration group(s) (group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 66** — Text Content — 3-item callout row with 2 icon(s) | Data/illustration slide — 3-item callout row with 2 icon(s), 5 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 66` to reuse this exact hand-designed composition for matching content in 'Text Content'. | title(idx0); subheadline(idx13); 3 standalone text box(es); 2 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 67** — Text Content — 5-item callout row with 4 icon(s) | Data/illustration slide — 5-item callout row with 4 icon(s), 9 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 67` to reuse this exact hand-designed composition for matching content in 'Text Content'. | title(idx0); subheadline(idx13); 5 standalone text box(es); 4 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_dark` |
| **Slide 68** — Text Content — 5-item callout row with 4 icon(s) [#68] | Data/illustration slide — 5-item callout row with 4 icon(s), 9 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 68` to reuse this exact hand-designed composition for matching content in 'Text Content'. | title(idx0); subheadline(idx13); 5 standalone text box(es); 4 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_dark` |
| **Slide 69** — 02 Collums (2-group diagram (0 text elements)) | Data/illustration slide — 2-group diagram (0 text elements), 5 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 69` to reuse this exact hand-designed composition for matching content in 'Text Content'. | title(idx0); subheadline(idx13); 5 standalone text box(es); 2 illustration group(s) (group of 0 text(s), group of 0 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 70** — 03 Collums (3-group diagram (0 text elements)) | Data/illustration slide — 3-group diagram (0 text elements), 7 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 70` to reuse this exact hand-designed composition for matching content in 'Text Content'. | title(idx0); subheadline(idx13); 7 standalone text box(es); 3 illustration group(s) (group of 0 text(s), group of 0 text(s), group of 0 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 71** — 04 Collums (4-group diagram (0 text elements)) | Data/illustration slide — 4-group diagram (0 text elements), 9 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 71` to reuse this exact hand-designed composition for matching content in 'Text Content'. | title(idx0); subheadline(idx13); 9 standalone text box(es); 4 illustration group(s) (group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 72** — Split Design (2-group diagram (8 text elements)) | Data/illustration slide — 2-group diagram (8 text elements), 10 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 72` to reuse this exact hand-designed composition for matching content in 'Text Content'. | title(idx0); subheadline(idx13); 2 standalone text box(es); 2 illustration group(s) (group of 4 text(s), group of 4 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split design` |
| **Slide 73** — Text Content — 1-group diagram (4 text elements) | Data/illustration slide — 1-group diagram (4 text elements), 6 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 73` to reuse this exact hand-designed composition for matching content in 'Text Content'. | title(idx0); subheadline(idx13); 2 standalone text box(es); 1 illustration group(s) (group of 4 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 74** — Text and images (simple text layout) | Simple content slide under 'Text and images' — clone via `clone_slide: 74`, or use the standard layout below directly for the same placeholder structure without the extra decoration. | title(idx0); subtitle(idx1); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 75** — Text and images — Text + 1 image layout | Data/illustration slide — text + 1 image layout, 3 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 75` to reuse this exact hand-designed composition for matching content in 'Text and images'. | title(idx0); subheadline(idx13); 2 standalone text box(es); 1 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 76** — Text and images (10-item text/number grid) | Data/illustration slide — 10-item text/number grid, 17 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 76` to reuse this exact hand-designed composition for matching content in 'Text and images'. | title(idx0); subheadline(idx13); 10 standalone text box(es); 7 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 77** — Text and images — 1-group diagram (0 text elements) | Data/illustration slide — 1-group diagram (0 text elements), 2 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 77` to reuse this exact hand-designed composition for matching content in 'Text and images'. | 2 standalone text box(es); 1 illustration group(s) (group of 0 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Full image Slide_dark_2` |
| **Slide 78** — Text and images — 3-item text layout | Data/illustration slide — 3-item text layout, 3 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 78` to reuse this exact hand-designed composition for matching content in 'Text and images'. | 3 standalone text box(es); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Full image Slide_dark_2` |
| **Slide 79** — Text and images — 3-group diagram (3 text elements) | Data/illustration slide — 3-group diagram (3 text elements), 5 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 79` to reuse this exact hand-designed composition for matching content in 'Text and images'. | 2 standalone text box(es); 3 illustration group(s) (group of 1 text(s), group of 1 text(s), group of 1 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Full image Slide_dark_2` |
| **Slide 80** — Text with image (4-item icon callout row (1 icons, 4 caption text elements)) | Data/illustration slide — 4-item icon callout row (1 icons, 4 caption text elements), 7 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 80` to reuse this exact hand-designed composition for matching content in 'Text and images'. | title(idx0); subheadline(idx13); 2 standalone text box(es); 4 illustration group(s) (group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s)); 1 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 81** — Text with image (4-item icon callout row (1 icons, 4 caption text elements)) [#81] | Data/illustration slide — 4-item icon callout row (1 icons, 4 caption text elements), 7 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 81` to reuse this exact hand-designed composition for matching content in 'Text and images'. | title(idx0); subheadline(idx13); 2 standalone text box(es); 4 illustration group(s) (group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s)); 1 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 82** — Text with horizontal image (4-item icon callout row (1 icons, 4 caption text elements)) | Data/illustration slide — 4-item icon callout row (1 icons, 4 caption text elements), 6 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 82` to reuse this exact hand-designed composition for matching content in 'Text and images'. | title(idx0); subheadline(idx13); 1 standalone text box(es); 4 illustration group(s) (group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s)); 1 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 83** — 02 Collums (2-item icon callout row (2 icons, 4 caption text elements)) | Data/illustration slide — 2-item icon callout row (2 icons, 4 caption text elements), 7 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 83` to reuse this exact hand-designed composition for matching content in 'Text and images'. | title(idx0); subheadline(idx13); 1 standalone text box(es); 2 illustration group(s) (group of 2 text(s), group of 2 text(s)); 2 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 84** — 03 Collums (8-item text/number grid) | Data/illustration slide — 8-item text/number grid, 11 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 84` to reuse this exact hand-designed composition for matching content in 'Text and images'. | title(idx0); subheadline(idx13); 8 standalone text box(es); 3 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 85** — 04 Collums (10-item text/number grid) | Data/illustration slide — 10-item text/number grid, 14 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 85` to reuse this exact hand-designed composition for matching content in 'Text and images'. | title(idx0); subheadline(idx13); 10 standalone text box(es); 4 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 86** — 03 Rows (2-item icon grid (3 icons, caption per icon)) | Data/illustration slide — 2-item icon grid (3 icons, caption per icon), 10 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 86` to reuse this exact hand-designed composition for matching content in 'Text and images'. | title(idx0); subheadline(idx13); 7 standalone text box(es); 2 illustration group(s) (group of 0 text(s), group of 0 text(s)); 3 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 87** — Table and Timeline Slides (simple text layout) | Simple content slide under 'Table Slides' — clone via `clone_slide: 87`, or use the standard layout below directly for the same placeholder structure without the extra decoration. | title(idx0); subtitle(idx1); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 88** — Table (simple text layout) | Simple content slide under 'Table Slides' — clone via `clone_slide: 88`, or use the standard layout below directly for the same placeholder structure without the extra decoration. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 89** — Table (simple text layout) [#89] | Simple content slide under 'Table Slides' — clone via `clone_slide: 89`, or use the standard layout below directly for the same placeholder structure without the extra decoration. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 90** — Timetable (4-group diagram (0 text elements)) | Data/illustration slide — 4-group diagram (0 text elements), 0 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 90` to reuse this exact hand-designed composition for matching content in 'Table Slides'. | title(idx0); subheadline(idx13); 4 illustration group(s) (group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 91** — Timetable (4-group diagram (4 text elements)) | Data/illustration slide — 4-group diagram (4 text elements), 4 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 91` to reuse this exact hand-designed composition for matching content in 'Table Slides'. | title(idx0); subheadline(idx13); 4 illustration group(s) (group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 92** — Timetable (22-group diagram (10 text elements)) | Data/illustration slide — 22-group diagram (10 text elements), 22 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 92` to reuse this exact hand-designed composition for matching content in 'Table Slides'. | title(idx0); subheadline(idx13); 12 standalone text box(es); 22 illustration group(s) (group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 93** — Timeline (7-group diagram (10 text elements)) | Data/illustration slide — 7-group diagram (10 text elements), 10 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 93` to reuse this exact hand-designed composition for matching content in 'Table Slides'. | title(idx0); subheadline(idx13); 7 illustration group(s) (group of 4 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 94** — Timeline (7-group diagram (0 text elements)) | Data/illustration slide — 7-group diagram (0 text elements), 14 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 94` to reuse this exact hand-designed composition for matching content in 'Table Slides'. | title(idx0); subheadline(idx13); 14 standalone text box(es); 7 illustration group(s) (group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 95** — Organigram Slides (simple text layout) | Simple content slide under 'Organigram Slides' — clone via `clone_slide: 95`, or use the standard layout below directly for the same placeholder structure without the extra decoration. | title(idx0); subtitle(idx1); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 96** — Organigram (connected multi-part illustration (28 text elements across 4 group(s))) | Data/illustration slide — connected multi-part illustration (28 text elements across 4 group(s)), 29 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 96` to reuse this exact hand-designed composition for matching content in 'Organigram Slides'. | title(idx0); subheadline(idx13); 1 standalone text box(es); 4 illustration group(s) (group of 7 text(s), group of 7 text(s), group of 7 text(s), group of 7 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 97** — Organigram (4-group diagram (16 text elements)) | Data/illustration slide — 4-group diagram (16 text elements), 18 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 97` to reuse this exact hand-designed composition for matching content in 'Organigram Slides'. | title(idx0); subheadline(idx13); 2 standalone text box(es); 4 illustration group(s) (group of 4 text(s), group of 4 text(s), group of 4 text(s), group of 4 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 98** — Leadership slide (17-item text/number grid) | Data/illustration slide — 17-item text/number grid, 32 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 98` to reuse this exact hand-designed composition for matching content in 'Organigram Slides'. | title(idx0); subheadline(idx13); 17 standalone text box(es); 15 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 99** — Organigram (15-item text/number grid) | Data/illustration slide — 15-item text/number grid, 15 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 99` to reuse this exact hand-designed composition for matching content in 'Organigram Slides'. | title(idx0); subheadline(idx13); 15 standalone text box(es); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 100** — Numbers and factsheets (simple text layout) | Simple content slide under 'Numbers and factsheets' — clone via `clone_slide: 100`, or use the standard layout below directly for the same placeholder structure without the extra decoration. | title(idx0); subtitle(idx1); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 101** — Text with numbers (4-group diagram (4 text elements)) | Data/illustration slide — 4-group diagram (4 text elements), 5 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 101` to reuse this exact hand-designed composition for matching content in 'Numbers and factsheets'. | title(idx0); subheadline(idx13); 1 standalone text box(es); 4 illustration group(s) (group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 102** — Text with numbers (3-group diagram (9 text elements)) | Data/illustration slide — 3-group diagram (9 text elements), 10 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 102` to reuse this exact hand-designed composition for matching content in 'Numbers and factsheets'. | title(idx0); subheadline(idx13); 1 standalone text box(es); 3 illustration group(s) (group of 3 text(s), group of 3 text(s), group of 3 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 103** — Factsheet page (8-group diagram (0 text elements)) | Data/illustration slide — 8-group diagram (0 text elements), 8 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 103` to reuse this exact hand-designed composition for matching content in 'Numbers and factsheets'. | title(idx0); subheadline(idx13); 8 standalone text box(es); 8 illustration group(s) (group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 104** — Diagrams and Charts (simple text layout) | Simple content slide under 'Diagrams and Charts' — clone via `clone_slide: 104`, or use the standard layout below directly for the same placeholder structure without the extra decoration. | title(idx0); subtitle(idx1); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 105** — Circle Segments (connected multi-part illustration (8 text elements across 1 group(s))) | Data/illustration slide — connected multi-part illustration (8 text elements across 1 group(s)), 12 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 105` to reuse this exact hand-designed composition for matching content in 'Diagrams and Charts'. | title(idx0); subheadline(idx13); 4 standalone text box(es); 1 illustration group(s) (group of 8 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 106** — Circle Segments (connected multi-part illustration (18 text elements across 2 group(s))) | Data/illustration slide — connected multi-part illustration (18 text elements across 2 group(s)), 18 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 106` to reuse this exact hand-designed composition for matching content in 'Diagrams and Charts'. | title(idx0); subheadline(idx13); 2 illustration group(s) (group of 6 text(s), group of 12 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 107** — Diagrams and Charts — 5-item text layout | Data/illustration slide — 5-item text layout, 5 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 107` to reuse this exact hand-designed composition for matching content in 'Diagrams and Charts'. | title(idx0); subheadline(idx13); 5 standalone text box(es); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 108** — Diagrams and Charts — 13-item text/number grid | Data/illustration slide — 13-item text/number grid, 13 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 108` to reuse this exact hand-designed composition for matching content in 'Diagrams and Charts'. | title(idx0); subheadline(idx13); 13 standalone text box(es); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 109** — Diagrams and Charts — 5-item icon callout row (6 icons, 8 caption text elements) | Data/illustration slide — 5-item icon callout row (6 icons, 8 caption text elements), 18 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 109` to reuse this exact hand-designed composition for matching content in 'Diagrams and Charts'. | title(idx0); subheadline(idx13); 4 standalone text box(es); 5 illustration group(s) (group of 2 text(s), group of 2 text(s), group of 2 text(s), group of 2 text(s), group of 0 text(s)); 6 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 110** — Diagrams and Charts — 5-item icon grid (2 icons, caption per icon) | Data/illustration slide — 5-item icon grid (2 icons, caption per icon), 10 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 110` to reuse this exact hand-designed composition for matching content in 'Diagrams and Charts'. | title(idx0); subheadline(idx13); 8 standalone text box(es); 5 illustration group(s) (group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s)); 2 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_dark` |
| **Slide 111** — Chart (5-group diagram (1 text elements)) | Data/illustration slide — 5-group diagram (1 text elements), 3 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 111` to reuse this exact hand-designed composition for matching content in 'Diagrams and Charts'. | title(idx0); subheadline(idx13); 2 standalone text box(es); 5 illustration group(s) (group of 1 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 112** — Keynumbers (4-group diagram (0 text elements)) | Data/illustration slide — 4-group diagram (0 text elements), 5 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 112` to reuse this exact hand-designed composition for matching content in 'Diagrams and Charts'. | title(idx0); subheadline(idx13); 5 standalone text box(es); 4 illustration group(s) (group of 0 text(s), group of 0 text(s), group of 0 text(s), group of 0 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 113** — Diagrams and Charts — 11-item text/number grid | Data/illustration slide — 11-item text/number grid, 11 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 113` to reuse this exact hand-designed composition for matching content in 'Diagrams and Charts'. | title(idx0); subheadline(idx13); 11 standalone text box(es); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 114** — Bar Chart (native chart) | Native chart example — model your own chart on it (chart type/axes/colors/data), don't clone the chart object itself (it's data-linked and won't carry your numbers). | title(idx0); subheadline(idx13); 1 standalone text box(es); 1 OLE object (think-cell data, not clonable — skipped automatically); 1 native chart; built on layout `Content Slide_empty_light` |
| **Slide 115** — Next Steps (simple text layout) | Simple content slide under 'Next Steps' — clone via `clone_slide: 115`, or use the standard layout below directly for the same placeholder structure without the extra decoration. | title(idx0); subtitle(idx1); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 116** — Next Steps (4-item callout row with 4 icon(s)) | Data/illustration slide — 4-item callout row with 4 icon(s), 8 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 116` to reuse this exact hand-designed composition for matching content in 'Next Steps'. | title(idx0); subheadline(idx13); 4 standalone text box(es); 4 picture(s); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 117** — Next Steps (1-group diagram (4 text elements)) | Data/illustration slide — 1-group diagram (4 text elements), 12 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 117` to reuse this exact hand-designed composition for matching content in 'Next Steps'. | title(idx0); subheadline(idx13); 8 standalone text box(es); 1 illustration group(s) (group of 4 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 118** — Next Steps (4-group diagram (4 text elements)) | Data/illustration slide — 4-group diagram (4 text elements), 8 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 118` to reuse this exact hand-designed composition for matching content in 'Next Steps'. | title(idx0); subheadline(idx13); 4 standalone text box(es); 4 illustration group(s) (group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 119** — Next Steps (6-group diagram (6 text elements)) | Data/illustration slide — 6-group diagram (6 text elements), 12 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 119` to reuse this exact hand-designed composition for matching content in 'Next Steps'. | title(idx0); subheadline(idx13); 6 standalone text box(es); 6 illustration group(s) (group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s), group of 1 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 120** — Next Steps (connected multi-part illustration (10 text elements across 1 group(s))) | Data/illustration slide — connected multi-part illustration (10 text elements across 1 group(s)), 10 fillable element(s) beyond the title and subheadline. Clone via `clone_slide: 120` to reuse this exact hand-designed composition for matching content in 'Next Steps'. | title(idx0); subheadline(idx13); 1 illustration group(s) (group of 10 text(s)); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Content Slide_empty_light` |
| **Slide 121** — Division Style/Theme — Simple text layout | Same structure as the standard layout below, re-skinned for Division Style/Theme — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subtitle(idx1); 1 standalone text box(es); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Title Page_dark` |
| **Slide 122** — Automotive Style — 4-item text layout | Same structure as the standard layout below, re-skinned for Automotive Style — use the standard layout directly and pair with the `automotive-*` icon set instead of cloning this one. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_dark` |
| **Slide 123** — Automotive Style — 4-item text layout [#123] | Same structure as the standard layout below, re-skinned for Automotive Style — use the standard layout directly and pair with the `automotive-*` icon set instead of cloning this one. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_light` |
| **Slide 124** — Automotive Style — Simple text layout | Same structure as the standard layout below, re-skinned for Automotive Style — use the standard layout directly and pair with the `automotive-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_dark` |
| **Slide 125** — Automotive Style — Simple text layout [#125] | Same structure as the standard layout below, re-skinned for Automotive Style — use the standard layout directly and pair with the `automotive-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_light` |
| **Slide 126** — Automotive Style — Simple text layout [#126] | Same structure as the standard layout below, re-skinned for Automotive Style — use the standard layout directly and pair with the `automotive-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_2` |
| **Slide 127** — Automotive Style — Simple text layout [#127] | Same structure as the standard layout below, re-skinned for Automotive Style — use the standard layout directly and pair with the `automotive-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_3` |
| **Slide 128** — Retail Style — 4-item text layout | Same structure as the standard layout below, re-skinned for Retail Style — use the standard layout directly and pair with the `retail-*` icon set instead of cloning this one. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_dark` |
| **Slide 129** — Retail Style — 4-item text layout [#129] | Same structure as the standard layout below, re-skinned for Retail Style — use the standard layout directly and pair with the `retail-*` icon set instead of cloning this one. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_light` |
| **Slide 130** — Retail Style — Simple text layout | Same structure as the standard layout below, re-skinned for Retail Style — use the standard layout directly and pair with the `retail-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_dark` |
| **Slide 131** — Retail Style — Simple text layout [#131] | Same structure as the standard layout below, re-skinned for Retail Style — use the standard layout directly and pair with the `retail-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_light` |
| **Slide 132** — Retail Style — Simple text layout [#132] | Same structure as the standard layout below, re-skinned for Retail Style — use the standard layout directly and pair with the `retail-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); built on layout `Split Content image_dark_1` |
| **Slide 133** — Retail Style — Simple text layout [#133] | Same structure as the standard layout below, re-skinned for Retail Style — use the standard layout directly and pair with the `retail-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_2` |
| **Slide 134** — Retail Style — Simple text layout [#134] | Same structure as the standard layout below, re-skinned for Retail Style — use the standard layout directly and pair with the `retail-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_3` |
| **Slide 135** — Real Estate Style — 4-item text layout | Same structure as the standard layout below, re-skinned for Real Estate Style — use the standard layout directly and pair with the `real-estate-*` icon set instead of cloning this one. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_dark` |
| **Slide 136** — Real Estate Style — 4-item text layout [#136] | Same structure as the standard layout below, re-skinned for Real Estate Style — use the standard layout directly and pair with the `real-estate-*` icon set instead of cloning this one. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_light` |
| **Slide 137** — Real Estate Style — Simple text layout | Same structure as the standard layout below, re-skinned for Real Estate Style — use the standard layout directly and pair with the `real-estate-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_dark` |
| **Slide 138** — Real Estate Style — Simple text layout [#138] | Same structure as the standard layout below, re-skinned for Real Estate Style — use the standard layout directly and pair with the `real-estate-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_light` |
| **Slide 139** — Real Estate Style — Simple text layout [#139] | Same structure as the standard layout below, re-skinned for Real Estate Style — use the standard layout directly and pair with the `real-estate-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); built on layout `Split Content image_dark_1` |
| **Slide 140** — Real Estate Style — Simple text layout [#140] | Same structure as the standard layout below, re-skinned for Real Estate Style — use the standard layout directly and pair with the `real-estate-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_2` |
| **Slide 141** — Real Estate Style — Simple text layout [#141] | Same structure as the standard layout below, re-skinned for Real Estate Style — use the standard layout directly and pair with the `real-estate-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_3` |
| **Slide 142** — Finance Style — 4-item text layout | Same structure as the standard layout below, re-skinned for Finance Style — use the standard layout directly and pair with the `financial-services-*` icon set instead of cloning this one. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_dark` |
| **Slide 143** — Finance Style — 4-item text layout [#143] | Same structure as the standard layout below, re-skinned for Finance Style — use the standard layout directly and pair with the `financial-services-*` icon set instead of cloning this one. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_light` |
| **Slide 144** — Finance Style — Simple text layout | Same structure as the standard layout below, re-skinned for Finance Style — use the standard layout directly and pair with the `financial-services-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_dark` |
| **Slide 145** — Finance Style — Simple text layout [#145] | Same structure as the standard layout below, re-skinned for Finance Style — use the standard layout directly and pair with the `financial-services-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_light` |
| **Slide 146** — Finance Style — Simple text layout [#146] | Same structure as the standard layout below, re-skinned for Finance Style — use the standard layout directly and pair with the `financial-services-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); built on layout `Split Content image_dark_1` |
| **Slide 147** — Finance Style — Simple text layout [#147] | Same structure as the standard layout below, re-skinned for Finance Style — use the standard layout directly and pair with the `financial-services-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_2` |
| **Slide 148** — Finance Style — Simple text layout [#148] | Same structure as the standard layout below, re-skinned for Finance Style — use the standard layout directly and pair with the `financial-services-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_3` |
| **Slide 149** — Health Style — 4-item text layout | Same structure as the standard layout below, re-skinned for Health Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_dark` |
| **Slide 150** — Health Style — 4-item text layout [#150] | Same structure as the standard layout below, re-skinned for Health Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_light` |
| **Slide 151** — Health Style — Simple text layout | Same structure as the standard layout below, re-skinned for Health Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_dark` |
| **Slide 152** — Health Style — Simple text layout [#152] | Same structure as the standard layout below, re-skinned for Health Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_light` |
| **Slide 153** — Health Style — Simple text layout [#153] | Same structure as the standard layout below, re-skinned for Health Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subheadline(idx13); built on layout `Split Content image_dark_1` |
| **Slide 154** — Health Style — Simple text layout [#154] | Same structure as the standard layout below, re-skinned for Health Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_2` |
| **Slide 155** — Health Style — Simple text layout [#155] | Same structure as the standard layout below, re-skinned for Health Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_3` |
| **Slide 156** — Education Style — 4-item text layout | Same structure as the standard layout below, re-skinned for Education Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_dark` |
| **Slide 157** — Education Style — 4-item text layout [#157] | Same structure as the standard layout below, re-skinned for Education Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_light` |
| **Slide 158** — Education Style — Simple text layout | Same structure as the standard layout below, re-skinned for Education Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_dark` |
| **Slide 159** — Education Style — Simple text layout [#159] | Same structure as the standard layout below, re-skinned for Education Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_light` |
| **Slide 160** — Education Style — Simple text layout [#160] | Same structure as the standard layout below, re-skinned for Education Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subheadline(idx13); built on layout `Split Content image_dark_1` |
| **Slide 161** — Education Style — Simple text layout [#161] | Same structure as the standard layout below, re-skinned for Education Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_3` |
| **Slide 162** — Education Style — Simple text layout [#162] | Same structure as the standard layout below, re-skinned for Education Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_2` |
| **Slide 163** — Corporate Style — 4-item text layout | Same structure as the standard layout below, re-skinned for Corporate Style — use the standard layout directly and pair with the `corporate-*` icon set instead of cloning this one. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_dark` |
| **Slide 164** — Corporate Style — 4-item text layout [#164] | Same structure as the standard layout below, re-skinned for Corporate Style — use the standard layout directly and pair with the `corporate-*` icon set instead of cloning this one. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_light` |
| **Slide 165** — Corporate Style — Simple text layout | Same structure as the standard layout below, re-skinned for Corporate Style — use the standard layout directly and pair with the `corporate-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_dark` |
| **Slide 166** — Corporate Style — Simple text layout [#166] | Same structure as the standard layout below, re-skinned for Corporate Style — use the standard layout directly and pair with the `corporate-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_light` |
| **Slide 167** — Corporate Style — Simple text layout [#167] | Same structure as the standard layout below, re-skinned for Corporate Style — use the standard layout directly and pair with the `corporate-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); built on layout `Split Content image_dark_1` |
| **Slide 168** — Corporate Style — Simple text layout [#168] | Same structure as the standard layout below, re-skinned for Corporate Style — use the standard layout directly and pair with the `corporate-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_2` |
| **Slide 169** — Corporate Style — Simple text layout [#169] | Same structure as the standard layout below, re-skinned for Corporate Style — use the standard layout directly and pair with the `corporate-*` icon set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_3` |
| **Slide 170** — Blue Style — 4-item text layout | Same structure as the standard layout below, re-skinned for Blue Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_dark` |
| **Slide 171** — Blue Style — 4-item text layout [#171] | Same structure as the standard layout below, re-skinned for Blue Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subtitle(idx1); 4 standalone text box(es); built on layout `Title Page with image_light` |
| **Slide 172** — Blue Style — Simple text layout | Same structure as the standard layout below, re-skinned for Blue Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_dark` |
| **Slide 173** — Blue Style — Simple text layout [#173] | Same structure as the standard layout below, re-skinned for Blue Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Chapter Slide_light` |
| **Slide 174** — Blue Style — Simple text layout [#174] | Same structure as the standard layout below, re-skinned for Blue Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subheadline(idx13); built on layout `Split Content image_dark_1` |
| **Slide 175** — Blue Style — Simple text layout [#175] | Same structure as the standard layout below, re-skinned for Blue Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_2` |
| **Slide 176** — Blue Style — Simple text layout [#176] | Same structure as the standard layout below, re-skinned for Blue Style — use the standard layout directly and use the generic `icons/` set instead of cloning this one. | title(idx0); subheadline(idx13); 1 OLE object (think-cell data, not clonable — skipped automatically); built on layout `Split Content image_dark_3` |
**Balance, not sparseness — avoid a slide that's 3 short bullets floating
in an otherwise empty dark rectangle.** A slide with only a handful of
short bullet points and nothing else leaves most of the slide as bare
background, which reads as unfinished rather than clean. When a slide's
content is genuinely a set of parallel points, **use a diagram (see
"Diagrams" section below) instead of a bare bullet list** — this is the
main lever available for the kind of well-composed, nothing-floating look
higher-end deck generators produce. This applies just as much on
multi-column layouts as on single-content ones: "Content Slide_2/3
columns" filled with plain short bullets in every column, and nothing
else on the slide, is exactly the same sparseness problem wearing a
different layout — convert each column to its own single-item `card_grid`
call instead (see "Diagrams" below). Reach for the 4-column icon layout
only when a diagram genuinely doesn't fit (e.g. the points need distinct
icons, not just titles/cards). It comes from choosing a composition whose
structure matches how much content the slide actually has, not from font
tricks.

**Icon captions on "Content Slide_4 columns" must be a phrase, not a bare
label.** The single most common way this deck ends up looking sparse in
practice: an icon paired with a one- or two-word caption like "Faster
Delivery" or "Security", with nothing else on the slide to fill the
space. Write each caption as a short phrase that says what it means or
why it matters — e.g. `"Cuts manual approval time in half"`, not
`"Faster Delivery"`. See `example_build_plan.json`'s 4-column slide for
the pattern to follow.

**The script checks all of these automatically and will warn (not
block) when they're violated — always read the warnings printed after
every `build`/`amend` run and fix the plan before sharing the file:**
- `check_sparse_captions` — flags any idx-15/17/19/21 caption on a
  "Content Slide_4 columns" layout under 4 words.
- `check_layout_rhythm` — flags a run of more than 2 consecutive slides
  using the same layout family (e.g. three "Content Slide_3 columns"
  slides in a row with no image-forward or chart slide breaking it up —
  this exact pattern is what prompted lowering the threshold from 3 to 2).
- `check_sparse_columns` — flags a "Content Slide_N columns" slide where
  every column is still plain bulleted text (never converted to a
  `card_grid`) with only a few short bullets each and nothing else on the
  slide — the multi-column version of the sparseness problem above. Also
  covers the single-column "Content Slide_1 column" case with its own
  tailored message (pair with an image, since that layout has no room to
  spread horizontally either).
- `check_diagram_variety` — flags more than 2 diagram-bearing slides in a
  row using the identical `diagram_type` (e.g. four `card_grid` slides
  back to back) — see the "Diagrams" section below for the full variety
  guidance this enforces.
- `check_title_subhead_collision` — flags a title that's likely to wrap
  to 2+ lines at its placeholder's actual width in a way that would crowd
  or overlap the subheading directly below it (see "Title length" below).
- `check_subhead_body_collision` — the same idea one level down: flags a
  subheading likely to wrap to 2+ lines in a way that would overlap the
  body content (bullets, a diagram, a chart) directly below it.
- `check_narrow_diagram_full_width_layout` — flags a `process_steps` or
  `accent_list` diagram placed on "Content Slide_1 column" (a full-width
  layout with no image slot), which leaves a large blank area since that
  content is naturally narrow — see "Pair narrow diagram types with an
  image" in the Diagrams section below.
- `check_empty_slide` — flags a slide with a title/subheading but no real
  body content at all (no filled bullets, image, chart, or diagram) —
  always a mistake (a dropped plan entry, a placeholder that never got
  filled), never a legitimate choice, so this fires every time. **This is
  also enforced automatically, not left to a warning that could be
  missed** — `build` and `amend` both strip any such slide out of the
  deck entirely before saving, printing a loud warning naming exactly
  which slide(s) and layout(s) were removed. A blank slide reaching the
  user is a bug with zero acceptable frequency — auto-removal guarantees
  that outcome unconditionally rather than depending on a warning being
  read and acted on. **This does not make the underlying gap OK to
  ignore** — the warning exists precisely so the planning mistake that
  produced a content-free slide (a dropped plan entry, a placeholder
  that never got filled) gets fixed for next time, not just silently
  swept away turn after turn. If a slide comes back removed, go find out
  why it had nothing on it and fix that, rather than treating the
  removal itself as the fix.
- `check_duplicate_content` — flags exact-duplicate text (4+ words)
  appearing in placeholder text on more than one slide. **For body
  content specifically (bullets, diagram descriptions — never title/
  subtitle, which can legitimately repeat), this is a hard block, not
  just a warning** — unlike the empty-slide case above, there's no safe
  automatic fix for duplicated content (removing the slide would also
  remove whatever legitimate, distinct part of it exists; removing just
  the duplicated text would leave a broken half-slide), so this refuses
  to save rather than guessing at a repair. This happened in practice at
  real severity: an amend that added new slides reused the identical
  3-bullet block verbatim across 9 different topically-distinct slides
  ("Store Audit App," "Inventory Request App," "Vehicle Inspection App,"
  etc. all showing the exact same "Reduce manual effort / Improve
  visibility / Enable mobile execution" bullets) instead of writing
  real, per-slide content — a generic filler pool substituted for actual
  per-slide writing. `build`/`amend` both refuse to save if this pattern
  is detected, raising `PLAN ERROR: BLOCKED: found body content repeated
  verbatim...` naming the text and every affected slide. **Never copy
  a slide's bullet/description text as a starting point for a new
  slide, even about a related topic — write each slide's content fresh
  for what that specific slide is about.** This applies with extra
  force when adding slides to an existing deck (see "Matching an
  existing deck's style" below) — the temptation to reuse the last
  slide's content as a template is exactly how this happened.
- `check_chapter_before_thank_you` — flags a "Chapter Slide" (section
  divider) sitting directly before the closing "Thank you and contact"
  slide with no real content slide in between — it introduces nothing
  and reads as filler, a different failure from `check_empty_slide`
  since the chapter slide typically has real title/tagline text, so it
  isn't "empty" by that check's definition. This happened in practice —
  see "Chapter/divider slides" below.
- `check_mixed_icon_usage` — flags an `accent_list` where some items got
  an `"icon"` and others didn't (e.g. item 1 and 3 have icons, item 2
  doesn't) — within one call, give every item an icon or none of them.
- `check_mixed_bullet_styles_on_slide` — flags a slide where different
  `"bullets"` placeholders use different `bullet_style` values from each
  other (e.g. one column uses `"check"`, the sibling column uses
  `"arrow"`) — pick one style per slide. Different slides can still use
  different styles freely.
- `check_chart_variety` — flags a deck with 2+ charts that are ALL the
  identical `chart_type` (e.g. every chart is a bar chart regardless of
  what the data actually shows) — see "Match chart type to the data"
  below.
- `check_all_charts_illustrative` — flags a deck where every chart is
  marked `"(Illustrative)"` — usually means real data was never actually
  searched for, not that it was genuinely unavailable. This happened in
  practice (a deck's only chart used fabricated numbers with no visible
  attempt at real research first). See "Charts" below.
- `check_generic_contact_info` — flags the "Thank you and contact"
  layout's contact block when it has no email address in it at all —
  almost always means it holds a generic description of the deck instead
  of real presenter details.

**Three fixes below are automatic — the script applies them itself after
every build/amend, you don't write anything for them, just know they're
happening:**
- **"Content Slide_4 columns" icon rows are auto-centered.** That layout
  positions its icon+caption row at a fixed spot sized only for the
  content itself, never stretched to fill the slide — left alone, a short
  caption set (the normal case) sits pinned near the top with a large
  empty area below it, exactly like an un-converted "Content Slide_N
  columns" bullet slide. The script re-centers the whole icon+caption
  block vertically in the space between the subheading and the slide's
  bottom margin.
- **Single-item `card_grid` siblings are auto-normalized to a uniform
  size.** The "one `card_grid` call per column" pattern (Retail/HR/
  Operations, one category per idx) sizes each card independently based
  on its own content — so if one column's text happens to wrap slightly
  differently than its neighbors, that card ends up a visibly different
  height and vertical position than the others, even though they're
  meant to read as a uniform row. This happened in practice. The script
  finds every single-item `card_grid` result on a slide and resizes them
  all to match the tallest one, at the same top position.
- **Every slide gets a slide number in the bottom-right corner,
  contrast-matched to that slide's theme.** The template's layouts
  define a slide-number placeholder there already (with the right
  color built in per dark/light layout), but python-pptx's `add_slide()`
  doesn't actually clone footer/slide-number/date placeholders onto new
  slides the way it does ordinary content placeholders — they were never
  present to begin with, which is why numbers never appeared even though
  the template looks like it supports them. The script draws the number
  directly instead: white text on dark-themed slides, navy on
  light-themed slides, same logic `_detect_theme` already uses for
  chart/diagram coloring. Idempotent across repeated amends — it redraws
  fresh each time rather than stacking duplicates as slides are added,
  removed, or reordered.

**Plain bulleted content (`"type": "bullets"`) is vertically centered
within its placeholder automatically** when it isn't the title/subheading
— a short list (2-4 items) in a tall placeholder used to sit pinned to
the top with a large dead zone below it; now it centers in the available
space. A long list that already fills the box looks identical either way,
so this is a pure improvement with no downside to watch for.

### Title AND subheading length — the script auto-shrinks to fit, but keep both short anyway

A title that wraps to 2 lines on a layout whose title box was sized for
exactly one line (with the subheading placeholder sitting right below,
zero slack) will visually crowd or overlap that subheading — this
happened in practice on a Split Content layout, where the title column is
only about half the slide's width, repeatedly, even after a warning was
added for it. The same thing happens one level down: a subheading that
wraps to 2 lines can overlap the body content (bullets, a diagram, a
chart) directly below it — also observed in practice on the same kind of
layout.

**The build script now fixes this automatically rather than only
warning about it.** When a single-line title or subheading is predicted
to wrap at its placeholder's actual width, `_set_text` shrinks that run's
font size just enough to fit on one line (calibrated from the template's
real default sizes — both title and subheading default to 28pt) — capped
at a floor of 20pt so it never shrinks so much it looks like a mistake in
its own right. Most of the time this is a barely-noticeable size
reduction that fully resolves the overlap with no plan changes needed.
`check_title_subhead_collision` and `check_subhead_body_collision` still
run afterward as a backstop for the rare case where even the floor size
doesn't fit (a genuinely very long title/subheading) — that's a real
"shorten this" signal, not a heuristic false alarm, since the shrink
already did everything it safely could.

This doesn't remove the value of keeping titles and subheadings
reasonably short in the first place — the flat 8-/15-word guidelines
elsewhere in this doc are still good defaults — it just means an
occasional longer one on a narrow layout no longer breaks the slide.

These warnings print after every `build`/`amend` run alongside the
existing duplicate-content and chart-coverage warnings. **Treat them the
same way: read them and fix the plan before sharing the file** — don't
generate, see a wall of caption warnings, and ship anyway.

See the layout table in the AFG PPT Builder system prompt for the full
list of layout names and their exact placeholder idx map (title,
subheadline, content columns, picture slots). Every `layout` value in the
JSON plan below must be one of the exact layout names from that table —
the script will error out with the full list of valid names if you pass
one that doesn't exist, rather than silently picking something close.

Two rules from that reference worth repeating here because they're the
most common mistake source:
- The subheadline placeholder (idx1 on Title Page layouts, idx13
  elsewhere) is **one short line only** — never bullets. Bulleted content
  goes in the `Content Placeholder` idx values (14/15/16/17/19/21
  depending on layout).
- `pic`/`clipArt` placeholders take an image path, never text. If you
  don't have a real image for a picture placeholder, leave that idx out
  of the plan entirely — the script will delete the unused placeholder
  rather than leaving it empty or filling it with a dummy image.

## Calling the script

**Build a new deck:**
```
python afg_deck_builder.py build --template AFG_template.pptx --plan plan.json --output <project_name>.pptx
```

**Amend an existing deck (always prefer this after the first build):**
```
python afg_deck_builder.py amend --input <project_name>.pptx --plan amend_plan.json --output <project_name>.pptx
```

**Inspect a deck's current placeholder-level content before amending:**
```
python afg_deck_builder.py inspect --input <project_name>.pptx
```
Use `inspect` whenever you're not certain what a slide currently contains
before writing an amend plan against it — don't guess from memory.

## Plan JSON format (build)

```json
{
  "slides": [
    {
      "layout": "Content Slide_2 columns_dark",
      "placeholders": {
        "0":  {"type": "text",    "value": "Why Power Apps Matters"},
        "13": {"type": "text",    "value": "Strategic value drivers"},
        "14": {"type": "bullets", "value": ["Build apps in weeks, not months", "Reduce manual spreadsheet work"]},
        "15": {"type": "bullets", "value": ["Lower total development cost", "Free up IT for higher-value work"]}
      }
    },
    {
      "clone_slide": 105,
      "placeholders": {
        "0":  {"type": "text", "value": "Rollout Phases"},
        "13": {"type": "text", "value": "Four stages to full adoption"}
      },
      "text_replacements": ["...", "01", "...", "02", "...", "03", "...", "04"]
    }
  ]
}
```

Two shapes of slide entry:
- **A normal slide**: `"layout"` + `"placeholders"`, as above.
- **A whole-cloned reference slide**: `"clone_slide"` (the reference
  slide's number in the original template) + `"placeholders"` (fills
  ONLY that slide's own title/subheadline idx) + `"text_replacements"`
  (fills everything else cloned from the reference slide — see "Cloned
  visuals" below for the full mechanics and the exact replacement
  ordering rule). Don't include `"layout"` on a `clone_slide` entry — the
  layout is determined by the reference slide itself.

- Every idx you want filled must be listed. Every idx that exists on the
  layout but isn't listed gets deleted automatically — you don't need to
  explicitly mark things for deletion, just omit them.
- `"type": "text"` → a single line (title, subheadline, or a picture's
  caption-free label). `"type": "bullets"` → a list of separate bullet
  lines in one content placeholder. `"type": "image"` → a file path to an
  image; only valid for `pic`/`clipArt` placeholders. `"type": "chart"` →
  a native, editable bar/line/pie chart that replaces the placeholder at
  its exact position/size (see Charts section below). `"type": "diagram"`
  → a native-shape illustration (numbered step chain or card grid) that
  replaces the placeholder at its exact position/size (see Diagrams
  section below). `"type": "cloned_visual"` → one shape/group cloned from
  a reference slide into this placeholder (see "Cloned visuals" below —
  this is placeholder-level; `clone_slide` above is slide-level and
  clones everything, not just one shape). A chart, diagram, or
  cloned_visual can replace any placeholder — text or picture — since all
  three fully substitute the shape rather than filling into it.
- Content must be written specifically for that slide's topic. Never reuse
  the same subheadline or bullet text verbatim across multiple slides —
  for body content this is a hard block, not just a warning (see "After
  running the script" below) — never copy one slide's bullets/description
  as the starting point for another, even about a related topic.
- `"type": "bullets"` is formatted automatically — the script applies a
  visible bullet marker, a larger font size than the template's inherited
  default, tighter line spacing between bullets, and extra space before
  the first bullet so it doesn't sit flush against the subheading above
  it. You don't need to add bullet characters or blank lines yourself;
  just give plain sentence strings in the list.
- Vary the bullet marker with an optional `"bullet_style"` alongside
  `"type": "bullets"` — e.g. `{"type": "bullets", "bullet_style": "arrow",
  "value": [...]}`. Options: `"dot"` (•, the default — plain content),
  `"arrow"` (→, for a sequence of related actions/capabilities — mirrors
  a style Gamma's own output uses), `"check"` (✓, for completed/included
  items), `"dash"` (–, a plainer, quieter list). This is glyph choice
  only; size, color, and spacing stay governed by the formatting above —
  it never turns into a manual override. An unrecognized value falls back
  to `"dot"` with a warning. **This has worked well in practice — actively
  reach for `"arrow"` or `"check"` instead of defaulting to `"dot"` every
  time**, the same way Section "Diagrams" asks you to vary diagram types:
  a deck where every single bullet list uses a plain dot is a missed,
  easy opportunity for variety, not a neutral default. **But vary it
  across slides, not within one** — if a slide has two bulleted columns
  (e.g. a 2-column layout), both must use the SAME `bullet_style`; a
  slide with one column in `"check"` and the other in `"arrow"` reads as
  inconsistent formatting, not a design choice. This happened in
  practice; `check_mixed_bullet_styles_on_slide` flags it. Pick one style
  per slide, vary the choice slide to slide across the deck.

## Cloned visuals — PRIORITY over procedural diagrams whenever a match exists

The template's 176-slide reference library isn't just a browsing aid for
picking a *layout* (see the "Reference library" note above) — many of
those slides contain a fully hand-designed illustration (a segmented
donut, a quarter-by-quarter timeline, a stat cascade, a process chevron
chain) built from ordinary native PowerPoint shapes. **Before building
any slide, check `reference_library_index.md` for a row whose
section/title matches that slide's actual content — this is a step in
the normal workflow, not something to reach for only after a slide
already looks repetitive.** Whenever a match exists, clone the real
thing and swap its text — do not redraw an approximation of it with the
procedural diagram engine below. A clone is pixel-exact to the designed
original; a procedural diagram is this skill's own estimate of it. The
procedural diagrams (`process_steps`/`card_grid`/`accent_list`/
`stat_row`), a plain layout, and bullets written from scratch remain
fully supported — but only as the fallback for content that has no
genuine match in the reference library, not as the default reached for
before checking.

**There are two ways to clone, and they're for different situations —
know which one you need before reaching for either:**

- **`clone_slide` (whole slide) — the default choice, and the one that
  genuinely lets you use any of the 176 reference slides directly,
  editing their real content in place.** Clones every non-placeholder
  shape on the reference slide at its ORIGINAL position — nothing is
  rescaled or refit, so multi-piece illustrations (several separate
  callouts around a central shape, a chain of connected badges) stay
  exactly as designed. The new slide uses that reference slide's own
  underlying layout, so its title/subheadline are filled the normal way
  through `placeholders`, alongside the cloned decoration. **This is
  also strictly more reliable than `cloned_visual` for multi-piece
  illustrations** — see the verified slide 120 example below, which is
  broken under single-shape cloning (rescaling distorts it) but renders
  perfectly under `clone_slide` (no rescaling happens at all).
- **`cloned_visual` (single shape) — for lifting ONE self-contained
  illustration out of a reference slide into a placeholder on a
  DIFFERENT layout**, e.g. putting just the circle-segments group into
  one side of a Split Content layout that also has an image on the other
  side. This is the right tool specifically when you need the
  illustration to live somewhere other than its own reference slide's
  layout, and are OK with it being scaled/centered to fit that
  destination placeholder.

If you're not sure which you need: if you want to use the reference
slide basically as it is (its own layout, its own composition), use
`clone_slide`. If you want to relocate just one piece of it onto a
slide built from other content, use `cloned_visual`.

### Workflow: catalog, then clone

**1. Find a candidate reference slide.** You generally already know
roughly where to look from the section it's likely filed under (see the
"Reference library" note's section list — e.g. a hierarchy/segmented
concept is probably in "Diagrams and Charts", a phased plan in "Next
Steps", a KPI-style summary in "Numbers and factsheets"). If you don't
know the exact slide number, search the template's text content for a
recognizable label (most reference slides carry their own descriptive
title as their first shape, e.g. "Circle Segments", "Timeline",
"Text with numbers") to find the slide index.

**2. Catalog it** to see every shape's id and, for groups, the exact
order `text_replacements` will be matched against:
```
python afg_deck_builder.py catalog --template AFG_template.pptx --slide 105
```
This prints every top-level shape's `shape_id` and type (flagging any
`OLE (not clonable)` shapes explicitly — these are think-cell data
objects and can't be cloned; the actual visual is almost always built
from separate, ordinary shapes elsewhere on the same slide, which *are*
clonable), plus, for each group, a numbered list showing each text box's
current placeholder content. **For `clone_slide`, the full replacement
order is every standalone (non-grouped) text box first in the order
`catalog` lists them, THEN each group's own internal order, group by
group in listed order** — read straight down `catalog`'s output
top-to-bottom to get this right; don't assume only the group's texts
matter, `clone_slide` picks up everything non-placeholder on the slide.

**3a. Clone the whole slide** (the usual choice — see slide 105's full
composition below, which includes standalone corner callouts `catalog`
would otherwise be easy to miss if you only looked at the group):
```json
{
  "clone_slide": 105,
  "placeholders": {
    "0": {"type": "text", "value": "Rollout Phases"},
    "13": {"type": "text", "value": "Four stages to full adoption"}
  },
  "text_replacements": [
    "Discovery phase covers stakeholder alignment",
    "Pilot phase covers early adopter feedback",
    "Scale phase covers business unit rollout",
    "Govern phase covers ongoing optimization",
    "Discover priority use cases", "01",
    "Scale across business units", "03",
    "Pilot with early adopters", "02",
    "Govern and optimize", "04"
  ]
}
```
This is a full slide-level entry — same level as a normal `{"layout":
..., "placeholders": ...}` entry in `build`'s `"slides"` list or
`amend`'s `"append_slides"` list, not something nested inside a
placeholder. `placeholders` here fills ONLY the title/subheadline (idx0/
idx13, which still exist as real placeholders on the cloned slide) — the
rest of the content comes entirely from `text_replacements` against the
cloned shapes. Verified end-to-end (built and visually inspected): the
four standalone corner labels plus the four-segment ring both render
correctly, matching the source slide's full design exactly.

**3b. Or clone just one shape into a placeholder** (when you need it
relocated onto a different layout):
```json
"14": {
  "type": "cloned_visual",
  "source_slide": 105,
  "source_shape_id": 9,
  "text_replacements": [
    "Executive sponsorship secured across all business units", "01",
    "Governance framework approved by IT leadership", "02",
    "Retail and HR pilots launched with early wins", "03",
    "Group-wide rollout scheduled for Q3", "04"
  ]
}
```
- `source_slide`: the reference slide's number in the **original**
  `AFG_template.pptx` (1-based, counting the template as shipped, before
  any slides are cleared for output).
- `source_shape_id`: the specific shape/group to clone, from `catalog`'s
  output — almost always the big `GROUP` shape that holds the actual
  illustration, not the slide's title/subheadline placeholders (those are
  filled normally via the slide's own idx system, same as any other
  slide).
- `text_replacements`: **supply one entry for every text box `catalog`
  listed, in that exact order** — including ones you don't want to
  change (like the "01"/"02"/"03"/"04" segment numbers above), since any
  box left unmatched keeps the template's original placeholder text
  (usually visible "Lorem ipsum..." copy, which must never ship). The
  script warns if you supply fewer replacements than there are boxes —
  treat that warning as a hard blocker, not a nice-to-fix.
- The cloned visual is automatically scaled (preserving its aspect
  ratio) and centered to fit whichever placeholder idx you target, so it
  lines up with the layout's grid — you don't need to compute position
  yourself.
- Text boxes inside the clone auto-shrink to fit their original
  box size if your replacement text is longer than the template's
  placeholder copy — but keep replacements reasonably close in length to
  what `catalog` showed regardless; these are fixed-size boxes from a
  hand-designed slide, not the auto-growing placeholders used elsewhere
  in this system, and very long text will just shrink to a barely-legible
  size rather than reflowing the layout.

### What can and can't be cloned

- **Can:** any ordinary shape or group — freeform illustrations,
  autoshapes, text boxes, pictures (image relationships are automatically
  rewritten so the copied picture still renders correctly in the output
  file).
- **Can't:** embedded OLE objects (think-cell data links) — `catalog`
  flags these explicitly, and the script refuses to clone one if asked.
  In every case checked so far, the actual visual a think-cell-linked
  slide displays is built from ordinary shapes sitting alongside (not
  inside) the OLE object, so this essentially never blocks cloning the
  illustration itself — just don't target the OLE shape's own id.

### Verified examples — tested end-to-end, safe to use as-is

Everything below was actually run through `build` and visually inspected
in the rendered output, not just cataloged. Reach for these first before
spending time cataloging your own candidate.

**Don't let "verified" become "the only one used."** This happened in
practice — after slide 105 was documented here, it became the *only*
cloned visual reached for across an entire deck, used repeatedly (or
exclusively) instead of one option among several. Cloning is one
technique among many for phased/proportional content, not a replacement
for variety:
- **Alternate between slide 105 (four segments) and slide 106 (six
  segments)** rather than defaulting to whichever one was documented or
  used first — pick based on how many segments the content actually
  has, not habit.
- **Most decks should mix cloned visuals with the procedural diagram
  types** (`process_steps`, `card_grid`, `accent_list`, `stat_row`) and
  plain content — cloning a hand-designed illustration is for the
  specific slide(s) where a phased/proportional shape is the best fit,
  not the default treatment for every diagram-shaped slide in the deck.
  A deck where every "diagrams" slide is the same cloned circle is a
  different flavor of the same repetition problem this section exists
  to fix.

- **Slide 105 — "Circle Segments" (four segments). Verified via BOTH
  modes.** As `clone_slide` (the fuller, recommended version): 4
  standalone corner callouts + the 4-segment ring, 12 text replacements
  in `catalog` order (4 standalone boxes first, then the ring group's 8:
  description/number ×4, clockwise from top). As `cloned_visual` (just
  the ring, shape_id 9, for relocating onto a different layout): 8
  replacements, description/number ×4. Both confirmed clean in testing —
  see the worked examples above/below.
- **Slide 106, shape_id 75 — "Circle Segments" (six segments).** Same
  pattern as 105's ring, 6 segments, 12 replacements (description, number
  ×6) for the `cloned_visual` form. Slide 106 also has a second group
  (shape_id 76, six description-only captions below the ring) — for
  `clone_slide`, both are picked up automatically; for `cloned_visual`,
  shape_id 75 alone is sufficient and is what was tested.
- **Slide 120 — "Next Steps" (five steps). Verified via `clone_slide`
  ONLY — do not use `cloned_visual`/single-shape mode for this one.**
  This is the clearest illustration of why the two modes exist: as
  `cloned_visual` (rescaling shape_id 74 to fit a different placeholder),
  it renders BROKEN — badges misaligned, one missing entirely, because
  rescaling distorts the group's internal child-shape proportions. As
  `clone_slide` (no rescaling — the whole slide, kept at its original
  size on its own native layout), it renders perfectly: all five
  connected badges correctly positioned and labeled. **This means the
  earlier "Next Steps slides aren't reliably cloneable" finding was
  specific to single-shape/rescaled cloning, not cloning in general** —
  `clone_slide` reopens the whole "Next Steps" section (115-120) as a
  legitimate option for phased/sequential content, not just 105/106's
  circle segments. 10 text replacements in `catalog` order: the group's
  own internal order for the 5 numbers (check `catalog` — it is NOT
  simply 01/02/03/04/05 in listed order, e.g. slide 120 lists
  01/02/04/03/05), then the 5 descriptions in that same positional order.
  **Get the number-to-description pairing from `catalog`'s exact order,
  don't assume sequential** — this is the one detail that's easy to get
  backwards (a test build with a naive sequential guess paired numbers
  and descriptions incorrectly; the shapes still rendered correctly
  positioned, just with mismatched pairs, which is entirely a
  plan-authoring mistake, not a rendering defect).

```json
"14": {
  "type": "cloned_visual",
  "source_slide": 105,
  "source_shape_id": 9,
  "text_replacements": [
    "Executive sponsorship secured across all business units", "01",
    "Governance framework approved by IT leadership", "02",
    "Retail and HR pilots launched with early wins", "03",
    "Group-wide rollout scheduled for Q3", "04"
  ]
}
```

**A naturally circular/square illustration like this one scales
by the SHORTER dimension to preserve its shape** — on a wide, short
placeholder ("Content Slide_1 column"), that leaves empty margin on
both sides rather than stretching the circle into an oval, which is
correct behavior, not a bug. If that empty margin looks like too much,
pair it with a narrower placeholder instead (one side of a Split
Content layout) — the same "match narrow content to a narrow layout"
principle that already applies to `process_steps`/`accent_list`. This
scaling behavior only applies to `cloned_visual` — `clone_slide` never
rescales at all, so it isn't a concern there.

**Slides 116-119 still aren't cleanly usable via EITHER mode** — they
aren't single groups (the numbers, icons, and text are separate
top-level shapes with no wrapping group), so there's no one
`source_shape_id` for `cloned_visual` to target, and even under
`clone_slide` (which picks up every top-level shape) they haven't been
individually verified the way 105/106/120 have. `catalog` one of these,
build it, and actually render it before trusting the result — or default
to the procedural `process_steps` diagram, which is reliable and already
AFG-styled, for next-steps content that doesn't map to slide 120's
specific five-step design.

**This is the general rule for cloning anything not on the verified list
above: catalog it, clone it, build the file, and actually look at the
rendered slide before including it in what you show the user.** A
cataloged shape list can look complete and still scale badly under
`cloned_visual`, or pair text incorrectly if you guess at ordering
instead of reading `catalog`'s exact sequence — the failure only shows
up in the render, not in `catalog`'s text output.

### Amending a slide that has a cloned visual

`amend`'s `--reference-template` flag must point at the **original**
`AFG_template.pptx`, not the working deck being amended — the working
deck already had its reference library stripped when it was first built,
so it has nothing left to clone from. Re-cloning to change a cloned
visual's text works the same way as replacing a chart or diagram at that
idx (see the amend plan format below).

## Diagrams — turn parallel points into an arranged illustration, but vary it, and don't illustrate everything

**Use a diagram whenever content is a set of parallel points, steps,
categories, or pillars, instead of reaching for `"bullets"`** — this is
the single biggest lever for making a deck look designed rather than
templated. A plain bulleted list is still the right call for prose-like
content that doesn't reduce to short parallel items.

**But not every slide should become one, and they should not all use the
same diagram type.** Two failure modes matter equally here:
- **Illustrating every single slide** — if the deck's story has a mix of
  parallel-list content and more narrative/explanatory content, leave the
  narrative slides as plain text/bullets. A deck where literally every
  slide is a diagram reads as busy and loses the contrast that makes the
  diagram slides land. Deliberately vary the *treatment*, not just the
  diagram type: some slides are plain text, some are a diagram, some are
  an image, some are a chart.
- **Using the same diagram type repeatedly** — four `card_grid` slides in
  a row (even if each one is individually correct) reads as "the same box
  shape over and over," not as a designed deck. **Rotate through the four
  diagram types below across the deck** the same way Section 2 already
  asks you to rotate layout shapes. The build script's
  `check_diagram_variety` warns automatically when more than 2
  diagram-bearing slides in a row use the identical `diagram_type` —
  treat that warning as something to fix, not background noise.

### The four diagram types — genuinely different visual languages, not variations on a box

All built from real, editable PowerPoint shapes in AFG brand colors —
never images. Pick based on both the content shape AND on what the
nearby slides already used:

- **`process_steps`** — vertical chain of numbered circular badges, each
  paired with a bold title + description. Use for sequences, workflows,
  "how it works," ordered roadmaps. Fills the full placeholder height by
  design (evenly divides it across items). Needs 2+ items — with 1 item,
  "full height / 1" stretches a single badge into an oddly tall shape;
  for a single point, use plain text instead.
- **`card_grid`** — one or more filled rounded-rect cards, each with a
  title + `"description"` or `"bullets"`. Use for parallel, non-sequential
  categories. The only one of the four that uses a boxed container — don't
  let it become the default just because it was the first one built.
  Automatically scales font size/padding to use more of the available
  space when the placeholder has room to spare (capped, and width-aware so
  narrow columns don't over-scale and break words). **The only one of the
  four where 1 item is valid** — its sizing is content-derived, not
  "stretch to fill height," so it's the right tool for the
  one-category-per-already-narrow-column pattern (see below).
- **`accent_list`** — vertical list, no box: each item gets either a
  small bundled icon (`"icon"`, e.g. `"icons/dark/speed.png"` on a dark
  slide) or, if no icon is given, a thin colored accent bar. A
  lighter-weight alternative to `card_grid` for feature/capability lists
  — reach for this when a nearby slide already used `card_grid` or
  `process_steps` and you want a different visual weight, or when each
  point genuinely has a distinct,
  relevant icon worth calling out (don't add icons just to add them — a
  plain accent bar is the right default when no icon is clearly relevant).
  Fills the full placeholder height by design, same as `process_steps` —
  **also needs 2+ items for the same reason**: a single item stretches
  into a tall, mostly-empty bar with the text floating in the middle
  (this happened in practice). Use single-item `card_grid` instead for
  one category per already-narrow column.
- **`stat_row`** — a horizontal row of large typographic numbers, a thin
  accent underline, and a label beneath — no boxes, no chart. Use for
  illustrative headline metrics that don't need a full chart (e.g. "70%
  faster," "3x adoption"). Distinctly different from all three above —
  good for a "business impact" style slide as a change of pace. Also the
  only one of the four besides multi-item `card_grid` that's genuinely
  full-width-friendly (see the layout-pairing rule right below).

### Pair narrow diagram types with an image — never strand them on a full-width layout

`process_steps` and `accent_list` are both naturally **narrow** content —
a numbered list or an accent list doesn't get better by being stretched
across the whole slide, so on "Content Slide_1 column" (a full-width
layout with no image slot at all), a large portion of the slide ends up
looking like empty background. This happened in practice. The fix is a
layout choice, not something to patch after the fact:

- Put narrow content (`process_steps`, `accent_list`, or a short plain
  bullet list) on a **Split Content image** layout instead, with a real
  photo/illustration filling the other side — mirrors Gamma's own use of
  this exact pattern (text column + a relevant illustration alongside).
  This is the single most valuable "make it look like Gamma" move
  available: a dynamic text-body-plus-illustration layout reads as
  considered where a lone narrow column on an otherwise-empty full-width
  slide reads as unfinished.
- If there's genuinely no image to use and the layout must be full-width,
  use a diagram type built for full width instead: multi-item `card_grid`
  (a horizontal row of cards spans the width naturally) or `stat_row`.
- The build script's `check_narrow_diagram_full_width_layout` flags this
  combination every time it occurs (not just past some repetition
  threshold, since there's no in-place fix) — treat it as a required
  change, not an FYI.

### The rule that matters most for `card_grid`: match the call to the placeholder's actual width

Two different situations look similar but need opposite handling — mixing
them up is what produced tiny, badly-wrapped, unreadable cards in
practice:

1. **One WIDE placeholder that should visually split into several cards**
   (e.g. idx14 alone on a "Content Slide_1 column" or one side of a Split
   Content layout — it spans a large single region). Here, make **one**
   `card_grid` call on that idx with **multiple items** — the script divides
   that one placeholder's width evenly across the cards.
2. **A layout that already gives each category its own separate,
   already-narrow placeholder** (e.g. "Content Slide_3 columns" → idx14,
   idx15, idx16 are three distinct placeholders, each already only
   roughly a third of the slide's width). Here, make **three separate**
   `card_grid` calls — **one per idx, each with exactly 1 item**. Never
   put a multi-item `card_grid` into just one of those idx values: cramming
   3 cards into a placeholder that's already only a third of the slide
   width divides it into ninths, and every title/bullet wraps onto 4+ lines
   of nearly illegible text. This is the single-item form of `card_grid` —
   it's not a special case, it's the normal way to turn "N columns" of bare
   bullets into "N cards", one call per existing column idx.

```json
"14": {"type": "diagram", "diagram_type": "card_grid", "items": [
  {"title": "Store Operations", "bullets": ["Store audits", "Visual merchandising checks", "Compliance reporting"]}
]},
"15": {"type": "diagram", "diagram_type": "card_grid", "items": [
  {"title": "Inventory & Stock", "bullets": ["Inventory requests", "Stock issue tracking", "Store support tickets"]}
]},
"16": {"type": "diagram", "diagram_type": "card_grid", "items": [
  {"title": "Customer Experience", "bullets": ["Customer feedback", "Service recovery actions", "Promotion tracking"]}
]}
```
This is the fix for the classic "Retail/HR/Operations Use Cases" slide
shape — 3 native template columns, each turned into its own titled card
with a mini bullet list, instead of 3 bare columns of plain text. See
`example_diagram_plan.json`'s third slide for the full worked version.

### Item schema

`process_steps`/`card_grid`/`accent_list` items: `{"title": str,
"description": str?, "bullets": [str]?, "icon": str? (accent_list only),
"number": int? (process_steps only)}`. `stat_row` items use a different,
simpler shape: `{"value": str, "label": str}` (e.g. `{"value": "70%",
"label": "Faster delivery"}`).

Rules:
- Give each `process_steps`/`card_grid`/`accent_list` item **either**
  `"description"` **or** `"bullets"`, never both — the script only
  renders one and warns if you supply both.
- `title` is required per item; a title with neither `description` nor
  `bullets` reproduces the exact "bare label" sparseness problem in shape
  form — the script warns on this.
- `stat_row` items need both `"value"` and `"label"` — the script warns
  if either is missing. **`"value"` must be a short number/metric** (e.g.
  `"70%"`, `"3x"`, `"15"`) — never a KPI name or word (`"Revenue"`,
  `"Margin"`). This happened in practice and, with 3+ items, produced
  visibly overlapping text between columns since a word is far wider than
  a short metric at the same font size. A named KPI with a category
  belongs in `card_grid` instead (`title`: the KPI name, `description`:
  the category) — `stat_row` is specifically for the number. The script
  warns when `"value"` looks like a word rather than a metric.
- `accent_list` items may add `"icon"`, a path to one of the bundled
  `icons/light/` or `icons/dark/` files matching the slide's theme (e.g.
  `"icons/light/governance.png"` on a light slide) — renders that icon in
  place of the plain accent bar. Only use it when the icon is genuinely
  relevant to that specific point (see the icons/ sourcing rules below);
  omit `"icon"` for a plain accent bar, which is the right default for
  most items. **Within one `accent_list` call, every item must have an
  icon, or none of them should** — giving 2 of 3 items an icon and
  leaving the third with a plain bar reads as a forgotten field, not a
  style choice. This happened in practice; the script's
  `check_mixed_icon_usage` flags it.
- `process_steps` and `accent_list` need **2+ items** — both distribute
  items evenly across the full placeholder height, so a single item
  stretches into an oddly tall, mostly-empty shape (script raises a hard
  error, not just a warning, since this always looks broken). `card_grid`
  and `stat_row` accept 1+ item.
- Up to 6 items per call. Above that, the script warns that items will be
  cramped — split across two slides instead of shrinking further.
- `number` is optional on `process_steps` items (defaults to 1-based
  position) — only set it explicitly if you need the displayed number to
  differ from list order.
- Colors, fonts, and spacing are fixed by the script from the AFG palette
  — you never choose them; just supply the item fields above.
- See `example_diagram_plan.json` for worked examples of all four types,
  including the one-`card_grid`-call-per-column fix for grouped/categorized
  columns, and an icon-per-item `accent_list` paired with an image.



## Charts — required at least once per deck, use real data or real research

You cannot generate images, but you CAN generate real, native, editable
charts — and every deck should include **at least one appropriate chart**
where the content supports it. Use a chart whenever the slide's content
is numeric (comparisons, trends, before/after figures, breakdowns, ROI,
timelines with numbers) rather than defaulting everything to bullet text.
The build/amend script warns you if a deck comes out with zero charts —
treat that warning as something to fix, not ignore, unless the deck
genuinely has no chartable content anywhere (rare).

**Actually search before defaulting to illustrative.** This happened in
practice: a deck's only chart used made-up "(Illustrative)" numbers with
no visible attempt at finding real ones first. The order of preference
in the schema below is not a suggestion — (1) real figures the user
provided, (2) real published data found via web search (industry
reports, vendor benchmarks, market research) with a `source` cited, and
only when neither is available, (3) an illustrative chart, clearly
labeled as such. Reaching step 3 should mean research was genuinely
exhausted, not skipped — the build script's `check_all_charts_illustrative`
flags a deck where every single chart is illustrative, which is the
signal to double check that.

```json
"14": {
  "type": "chart",
  "chart_type": "bar",
  "title": "Global Low-Code Market ($B)",
  "x_axis_title": "Year",
  "y_axis_title": "Market Size ($B)",
  "categories": ["2022", "2023", "2024", "2025"],
  "series": { "Market size": [13.8, 18.3, 23.9, 31.1] },
  "source": "Gartner, Low-Code Development Technologies Forecast"
}
```

- `chart_type`: `"bar"` (vertical columns), `"bar_horizontal"`, `"line"`,
  or `"pie"`. **Match the type to what the data is actually showing, and
  actively vary it across the deck** — the build script's
  `check_chart_variety` flags a deck where 2+ charts all use the same
  type, since that's usually habit (always reaching for bar) rather than
  a deliberate read of the data:
  - Trend over time (monthly/quarterly figures, growth over a period) →
    `"line"`.
  - Share of a whole (percentage breakdown that sums to ~100%, mix of
    categories) → `"pie"`.
  - Comparison across discrete categories (this region vs that region,
    this product vs that product) → `"bar"` or `"bar_horizontal"` (the
    latter when category labels are long).
  A deck with several charts that are all genuinely category comparisons
  can legitimately be all bar charts — the point is to check that against
  what each chart's data actually is, not to force variety for its own sake.
- `categories`: the x-axis / slice labels. `series`: one or more named
  series, each a list of numbers the same length as `categories` (pie
  charts should have exactly one series).
- `x_axis_title` / `y_axis_title` (optional but encouraged): axis labels.
  Omit them only when the axis is genuinely self-explanatory from the
  category names alone.
- `source` (required whenever the figures came from a web search rather
  than the user directly): a short citation string. The script renders it
  as a small italic caption beneath the chart automatically — don't also
  put the citation in the chart title or a separate text box.
- The chart is placed at the exact position/size of whichever placeholder
  idx you assign it to — pick a `Content Placeholder` idx (or a `pic` idx
  if you want it where an image would have gone) so it lines up with the
  template's grid instead of floating at an arbitrary position.
- **Colors, data labels, and axis text are all handled automatically** —
  the script detects whether the target slide is a `_dark` or `_light`
  layout and applies the matching AFG palette (navy/blue/grey series
  colors, white text on dark slides, navy text on light slides) with
  visible data labels on every bar/point/slice. You don't need to specify
  colors yourself.
- Pie chart data labels show the raw series values as given, not
  auto-converted to a `%` symbol — if you want the labels to read as
  percentages, make the series values already sum to ~100 and note in
  the chart title that units are %, e.g. `"Use Case Split (%)"`.

**Where the numbers come from, in priority order:**

1. **Real data the user gave you.** Use it directly, exactly as given —
   never round, invent, or "improve" numbers the user provided.
2. **Real data from the web**, when the user hasn't supplied figures.
   Search for a genuinely relevant, citable statistic for the slide's
   topic (industry benchmark, published market report, vendor-reported
   adoption numbers, etc.) rather than defaulting straight to invented
   placeholder numbers. Always set `source` to a short citation when you
   use this path, and mention in your reply to the user where the figure
   came from so they can verify it.
3. **Illustrative, clearly labeled, only if neither of the above is
   available** (e.g. no real data exists yet for a forward-looking
   projection and a relevant web figure can't be found). It's fine to
   build an illustrative chart to convey a shape or direction (e.g. "cost
   trending down over three phases"), but it must say "(Illustrative)" in
   the chart title itself (e.g. `"title": "Projected Cost Reduction
   (Illustrative)"`), and you must say so in your reply — never presented
   as if it were real data.

Do not source "graph images" from the web to represent data — a scraped
chart image shows someone else's finished chart, not this deck's content,
and can't be labeled, recolored, or corrected the way a native chart can.
Any graph in the deck should be a native chart built by this script: real
user data, real web-researched data with a citation, or clearly-marked
illustrative — never a found image standing in for a chart.


## Plan JSON format (amend)

```json
{
  "edits": [
    {
      "slide_index": 3,
      "placeholders": {
        "14": {"type": "bullets", "value": ["New first point", "New second point"]}
      },
      "delete_placeholders": [16]
    }
  ],
  "delete_slides": [7],
  "append_slides": [
    {
      "layout": "Thank you and contact_dark",
      "placeholders": { "...": "..." }
    },
    {
      "clone_slide": 105,
      "placeholders": { "0": {"type": "text", "value": "..."}, "13": {"type": "text", "value": "..."} },
      "text_replacements": ["...", "01", "...", "02", "...", "03", "...", "04"]
    }
  ]
}
```

`append_slides` entries can be either shape of slide entry from the
`build` format above — a normal `"layout"` + `"placeholders"` entry, or
a `"clone_slide"` + `"placeholders"` + `"text_replacements"` entry. Both
respect the same closing-slide reordering (`_ensure_closing_slide_is_last`)
and duplicate-content block as `build`.

`slide_index` is 1-based, matching how you should refer to slides when
talking to the user. Amend edits only touch the idx values you list — it
does not delete unmentioned placeholders on that slide the way `build`
does, since an amendment is a partial change, not a full rebuild of that
slide's content.

**When an amend changes a caption's topic, check whether its paired icon
still matches — amend won't do this for you.** Since amend only touches
the idx values explicitly listed, editing a 4-column caption's text (or
an `accent_list` item's text) without also updating its paired image/icon
idx leaves the old icon in place even though it no longer matches the new
topic (e.g. the caption changes from "HR onboarding" to "Finance
approvals" but the teamwork icon stays). Before finalizing an amend that
changes what a slide's content is *about*, re-check every icon on that
slide against its current caption and include an updated `"type":
"image"` entry for any idx whose icon no longer fits — don't assume
editing the text alone is a complete edit.

### The closing "Thank you and contact" slide needs real presenter details

idx16's contact block exists to say who to follow up with — it must hold
the actual presenter's name, title, organization, and email (e.g. `"Jane
Doe\nPower Platform Lead\nAFG Digital\njane.doe@afg.com"`), never a
description of the presentation itself (`"Executive Analytics
Presentation"` is not contact info — this happened in practice).

**Default assumption: this is the requesting user's own deck, so use
their own details unless told otherwise.** Don't ask "should I include
your name/email, or omit contact details?" — that's an unnecessary
back-and-forth for something that has an obvious default. Instead:
- If the user's name, role, or email is already known from earlier in
  the conversation, use it directly, no confirmation needed.
- If some of it is missing (e.g., name known but not email), fill in
  what's known and ask only for the specific missing piece, framed as
  gathering the last detail rather than seeking permission — e.g. "What
  email should go on the closing slide?" not "Would you like me to add
  your email?"
- Only skip the contact block entirely if the user explicitly says to
  omit it, or says the deck isn't theirs to be credited on (presenting
  on someone else's behalf, a template for others to reuse, etc.).
- Never invent a name, title, or email that wasn't given or known —
  asking for the missing piece is always better than fabricating one.

The build script's `check_generic_contact_info` flags this block when it
contains no email address at all, which is the simplest reliable signal
that it's not real contact info.

### Chapter/divider slides need a real section to introduce — never use one as filler

"Chapter Slide" (title + tagline + full-bleed image, no bullets/data/
diagram by design) exists to introduce a section of several content
slides that follow it — it is not a content slide itself and shouldn't
be asked to carry a message on its own. **Never place a Chapter Slide
immediately before the closing "Thank you and contact" slide** — this
happened in practice (a "2026-2030 Vision" chapter slide sitting right
before Thank You, with nothing after it to introduce). With nothing
following it, a chapter slide introduces nothing — it reads as a
filler slide with no real content, exactly the kind of unnecessary
slide to avoid. If there's a genuine closing thought to convey, fold it
into the Thank You slide's own tagline (idx13) instead of giving it a
whole separate divider slide. The build script's
`check_chapter_before_thank_you` flags this exact adjacency. More
generally: only use a Chapter Slide when at least one real content slide
follows it in the same section — if a "section" would only ever contain
the chapter slide itself, it doesn't need a divider, it needs to just be
folded into the surrounding content directly.

## Professional-polish checklist — the concrete "make it look like Gamma" list

Everything below is already implemented in this skill; this section is
just the checklist to actually reach for it while planning a deck, rather
than defaulting to the plainest option on every slide:

1. **Vary the treatment slide to slide.** Not every slide is a diagram;
   not every diagram is the same type. Mix plain text, `card_grid`,
   `process_steps`, `accent_list`, `stat_row`, charts, and images
   deliberately (Diagrams section above).
2. **Match content volume to layout capacity.** A handful of short
   bullets in a large layout is the #1 cause of a slide looking
   unfinished (Section 2A-style sparseness) — use a diagram instead, or a
   smaller layout.
3. **Pair narrow content (`process_steps`, `accent_list`, a short bullet
   list) with an image via a Split Content layout**, rather than
   stranding it on a full-width single-column layout with nothing on the
   other side — this dynamic text-plus-illustration composition is the
   single most valuable "make it look like Gamma" move available (see
   "Pair narrow diagram types with an image" above).
4. **Give every diagram item a real phrase, not a bare label** — "Cuts
   approval time in half," not "Faster." One- or two-word items are the
   single biggest tell of a rushed deck.
5. **Use a relevant icon per bullet where it genuinely adds meaning** —
   `accent_list`'s optional `"icon"` field, or the 4-column icon layout —
   and vary the bullet marker (`bullet_style`: arrow/check/dash) instead
   of defaulting to a plain dot on every single bulleted list. **Stay
   consistent within a slide** — every item in one `accent_list` gets an
   icon or none do; every bulleted column on one slide uses the same
   `bullet_style`. Vary the choice across slides, not within one.
6. **Keep titles AND subheadings short relative to their actual
   placeholder width** — especially on Split Content / narrower layouts,
   where a wrapped title or subheading can crowd the content below it
   (see "Title length" above).
7. **Include at least one real chart** with real data or cited research,
   **and match each chart's type to its data** — trend over time is a
   line chart, share of a whole is a pie chart, category comparison is a
   bar chart. An illustrative `stat_row` is not a substitute for a chart
   when the content is genuinely quantitative (Charts section below).
   **Actually search for real data before defaulting to illustrative** —
   reaching for "(Illustrative)" without trying should never be the
   default path (see "Actually search before defaulting" above).
8. **Every slide needs something substantive using its space** — an
   illustration/image, a chart, a diagram, or a KPI callout — not just
   a title and a thin column of text with the rest of the slide bare.
9. **The closing slide needs the real presenter's name, title, and
   email** — never a generic description of the deck (see "The closing
   Thank you and contact slide" above).
10. **Never place a Chapter/divider slide with nothing following it** —
   especially not directly before the closing Thank You slide (see
   "Chapter/divider slides" above). A section divider that introduces no
   section is filler.
11. **Read every warning the script prints and fix the plan before
   sharing the file** — every item above has a corresponding automated
   check; a clean warning list is the closest thing to a final QA pass
   this skill can do for you.
12. **Avoiding repetition isn't the same as using the whole catalog** —
   the rhythm check only flags runs of the *same* layout back to back; a
   deck can pass it while never once touching Quote, Chapter/divider,
   Split design, Full image, or bilingual, just by alternating between
   two or three "safe" layouts the whole way through. For any deck longer
   than ~6 slides, explicitly check whether the content actually calls
   for one of these: a stakeholder or customer quote → **Quote Slide**;
   a natural break between major sections → **Chapter Slide**; an
   explicit two-way comparison or before/after → **Split design**; one
   strong single visual moment (a launch, a location, a hero shot) →
   **Full image Slide**. These are real, fully-supported layouts, not
   edge cases — a deck that never reaches for any of them across an
   entire presentation is a sign they weren't considered, not that they
   never applied.
13. **A deck with zero `clone_slide` usage at all is a sign the 176-slide
   reference table wasn't actually checked per slide** — this happened
   in practice, repeatedly, including right after this table was added:
   a deck built using only the ~34 base layouts, with the far larger
   reference library never touched. Checking the reference table is a
   REQUIRED step for every slide in the outline (see "All 176 template
   slides" above), not something to consider only once something already
   looks repetitive. Before finalizing the outline, check each slide's
   content against that table directly: does any slide describe a
   phased plan, a proportional breakdown, parallel capability pillars, a
   sequence of numbered steps, an org chart, a KPI/factsheet layout, or
   any other shape a reference slide covers? If a row matches, use
   `clone_slide: <#>` for it rather than defaulting to a plain layout,
   `card_grid`, or bullets — a hand-designed illustration reads as more
   considered than this skill's own procedural approximation of the same
   idea. **The reverse also happened in practice — a deck using the
   SAME cloned slide (105) for every diagram-shaped slide, exclusively,
   with no procedural diagrams and no other reference slides at all.**
   Cloning is one option among many across all 176, not a replacement
   for the diagram-type variety this section already asks for, and not
   a reason to fixate on the two or three slide numbers that happen to
   be documented as pre-verified — alternate across the full table based
   on what each slide's content actually is, not on which examples were
   easiest to remember.

## After running the script — the warning list is a gate, not a suggestion

The script prints a slide-by-slide summary (layout used + title text) and
a warning list after every `build`/`amend` run. **Exit code 0 only means
the plan was structurally valid — right idx, right placeholder types. It
says nothing about whether the deck is well-composed.** The warning list
is what actually tells you that, and it exists specifically to catch
things like four consecutive identical layouts, zero charts, or no
diagram variety — the exact kind of output that looks technically fine
but reads as repetitive and unfinished.

**One category of problem doesn't wait for you to read a warning at
all.** A blank slide (no real body content anywhere on it) gets
automatically removed from the deck before the file is saved — the file
still gets written (unlike the duplicate-content case below, which does
block), but a loud warning names exactly which slide(s) and layout(s)
were stripped and why. **Seeing that warning still means the plan has a
gap that needs fixing** — the deck now has fewer slides than the plan
called for, which changes numbering and pacing, and whatever content was
supposed to be on that slide is simply gone. Don't treat the removal as
having handled it: go back to the plan, figure out why that slide ended
up with nothing on it (a dropped entry, a placeholder that never got
filled), and either add the real content back via `amend` or confirm the
slide genuinely wasn't needed. This is deliberately not something to
just note and move past — a blank slide never reaching the user is the
guarantee, not an excuse to skip finding out why one was about to.

**Before calling `present_files` or telling the user the deck is
ready:**
1. If the warning list is non-empty, revise the plan to address each one
   specifically and re-run. Repeat until it's empty.
2. The only acceptable alternative to fixing a warning is a specific,
   stated reason it doesn't apply (e.g. a genuine 3-slide update where
   nothing is numeric, so chart-coverage doesn't apply) — a conscious,
   explainable call, never silence.
3. This applies after every `amend` too, not just the first `build` — an
   amend can introduce the same rhythm/variety/coverage problems from
   scratch (e.g. amending three slides in a row to the same layout).

Summarize what was built or changed in your reply to the user in one line
per slide; don't paste the full slide text back as a wall of text.

## Sourcing images — two folders, two different jobs

There are two bundled folders, and they are **not interchangeable** —
using the wrong one in the wrong slot produces badly stretched, blown-up
visuals (this happened before: a small square icon was used to fill a
wide title banner and a tall portrait panel, and both looked broken).

### `icons/` — compact badges for the 4-column layout and `accent_list` icons ONLY

Each file is a simple flat line-art graphic, 400×400, **transparent
background** — every icon exists in two color variants, one per theme:

- `icons/light/<name>.png` — navy/blue, for use on light-themed slides
  (layout name ends `_light`).
- `icons/dark/<name>.png` — white/blue, for use on dark-themed slides
  (layout name ends `_dark`).

**Always pick the variant matching the slide's own theme, never mix
them.** A dark-variant (white) icon on a light slide is invisible; a
light-variant (navy) icon on a dark slide is invisible. This is a real
bug that existed before the transparent versions were introduced (opaque
white-background icons showed as a visible white box on dark slides) —
the fix depends entirely on using the theme-matched folder every time.

These are sized and composed for two specific uses: the small square
picture slots on **"Content Slide_4 columns_dark/light"**
(idx14/16/18/20, a badge next to one short text label), and the optional
`"icon"` field on `accent_list` diagram items (see "Diagrams" section).
They need no capability to be enabled; they're just files, always
available the moment this skill is active.

| file (under `icons/light/` or `icons/dark/`) | concept |
|---|---|
| `speed.png` | speed, fast delivery, agility |
| `cost.png` | cost reduction, savings, efficiency |
| `productivity.png` | productivity, output, performance |
| `growth.png` | growth, scaling, upward trend |
| `governance.png` | governance, control, compliance approval |
| `security.png` | security, data protection, access control |
| `checklist.png` | audits, checklists, process steps |
| `teamwork.png` | HR, people, collaboration, culture |
| `retail.png` | retail, stores, customer operations |
| `finance.png` | finance, procurement, budgets |
| `operations.png` | operations, facilities, maintenance |
| `technology.png` | technology, digital, platforms, apps |
| `data.png` | data, analytics, reporting, insights |
| `mobile.png` | mobile apps, field tools, self-service |
| `roadmap.png` | roadmaps, timelines, phased rollouts |
| `corporate-lamp.png` | office setup, workspace, desk |
| `corporate-phone.png` | contact, communication, front desk |
| `corporate-briefcase.png` | business, professional services, work |
| `corporate-clock.png` | time tracking, scheduling, deadlines |
| `corporate-monitor.png` | desktop work, reporting, office IT |
| `retail-storefront.png` | store operations, physical retail |
| `retail-cart.png` | shopping, checkout, e-commerce |
| `retail-tag.png` | pricing, discounts, promotions |
| `retail-gift.png` | gifting, loyalty rewards, promotions |
| `retail-delivery-truck.png` | delivery, logistics, fulfillment |
| `real-estate-house.png` | residential property, housing |
| `real-estate-bed.png` | furnished units, hospitality, interiors |
| `real-estate-key.png` | leasing, handover, access |
| `real-estate-for-rent.png` | listings, availability, rent/sale signage |
| `real-estate-building.png` | commercial property, multi-unit buildings |
| `automotive-car.png` | vehicles, fleet, dealership |
| `automotive-ev-charging.png` | EV charging, electric vehicles |
| `automotive-steering-wheel.png` | driving, aftersales, service |
| `automotive-warning.png` | alerts, diagnostics, safety |
| `automotive-battery.png` | vehicle battery, EV power |
| `financial-services-credit-card.png` | payments, cards, transactions |
| `financial-services-trend-chart.png` | growth trend, performance illustration |
| `financial-services-pie-chart.png` | portfolio mix, breakdowns, allocation |
| `financial-services-atm.png` | banking, cash services |
| `financial-services-wallet.png` | accounts, balances, personal finance |

Match the closest concept to the slide's topic. Reuse the same icon file
if the same concept recurs across slides — don't hunt for variety for
its own sake.

**Never use a file from `icons/` on any other layout's image slot** —
specifically not Title Page with image, Chapter Slide, Split Content
image, or Full image Slide. Those are large hero-style regions with very
different proportions (some over 3.5:1 wide, some tall portrait panels);
stretching a small square badge into one of them crops it down to an
unrecognizable sliver or blows it up into a blurry mess. The build script
enforces this as a hard rule (it checks the target placeholder's size and
refuses the plan if an `icons/` file targets a slot that large) — but
plan around it up front rather than relying on the error to catch it.
For those hero slots: use a real image from `images/` (below), or choose
a layout that doesn't need a large image at all.

**Adding more icons yourself:** create BOTH a `icons/light/<name>.png`
(navy/blue) and `icons/dark/<name>.png` (white/blue) version with a
transparent background, same ~400×400 stroke-icon style as the existing
set, and add one row to the table above (the filename is shared between
both folders, only the color differs). A single opaque-background icon
with no dark-theme counterpart will look broken the first time it's used
on a dark slide — this happened before across the whole original set.

### `images/` — real photos/illustrations for hero-style slots

A separate bundled folder for actual photography or illustration,
intended for the large single-image layouts (Title Page with image,
Chapter Slide, Split Content image, Full image Slide) that `icons/`
files must never go into. This folder starts empty — add real,
cleared-for-use images here as they become available (see
`images/README.md` inside the skill package for naming and licensing
notes). Reference them the same way as icons: `"type": "image", "value":
"images/<filename>"`.

This folder is also the intended drop-in point for future generated
images once an image generator is connected — once that's wired up,
generated images can be saved here (or referenced directly) without
changing how build/amend plans work.

### If neither folder has what a hero slot needs

Ask the user for a real image, or — only if enabled in this environment
and appropriate for the slot's actual proportions — use Copilot Studio's
Image Generator (Settings → Capabilities). Otherwise, choose a different
layout without a large image slot rather than forcing a mismatched
visual into place.

Rules that still apply regardless of source:
- Only use abstract/icon-style graphics from `icons/` — never fabricate a
  "photo" purporting to show real AFG people, offices, or products;
  reserve photographic content for real images in `images/` or
  user-supplied uploads.
- If nothing suitable is available, leave that `pic` idx out of the plan
  so the script deletes the placeholder — never fall back to a
  blank/solid-color filler image, and never force a wrong-shaped image
  into a slot just to avoid leaving it empty.
