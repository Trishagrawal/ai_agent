#!/usr/bin/env python3
"""One-off generator for reference_library_index.md, run against the updated
AFG_template.pptx. Not part of the shipped skill -- deleted after use."""
import re
import sys
from pptx import Presentation
from pptx.oxml.ns import qn

TEMPLATE = "skills_extracted/AFG_template.pptx"
OUT = "skills_extracted/reference_library_index.md"

PLACEHOLDER_TITLES = (
    "click to add", "please insert", "headline", "lorem ipsum", "text",
    "al-futtaim template", "logo", "logos",
)


def clean(t):
    if t is None:
        return ""
    return re.sub(r"[\x0b\x0c\n]", " ", t).strip()


def is_placeholderish(title):
    low = title.lower()
    if not low:
        return True
    if low.isdigit():
        return True
    for p in PLACEHOLDER_TITLES:
        if low.startswith(p):
            return True
    return False


def count_group_texts(shape):
    n = 0
    if shape.shape_type == 6:  # GROUP
        for sub in shape.shapes:
            n += count_group_texts(sub)
    elif shape.has_text_frame and shape.text_frame.text.strip():
        n += 1
    return n


def analyze(slide):
    title = ""
    n_title = n_subtitle = n_subhead = 0
    other_ph = []
    standalone_text = 0
    groups = []          # list of leaf-text counts
    pictures = 0
    ole = 0
    for sh in slide.shapes:
        if sh.is_placeholder:
            idx = sh.placeholder_format.idx
            if idx == 0:
                n_title += 1
            elif idx == 1:
                n_subtitle += 1
            elif idx == 13:
                n_subhead += 1
            else:
                other_ph.append(idx)
            if sh.has_text_frame and sh.text_frame.text.strip() and not title:
                title = clean(sh.text_frame.text)
            continue
        st = sh.shape_type
        if st == 7:  # OLE
            ole += 1
        elif st == 6:  # GROUP
            groups.append(count_group_texts(sh))
        elif st is not None and "PICTURE" in str(st):
            pictures += 1
        elif sh.has_text_frame and sh.text_frame.text.strip():
            standalone_text += 1
            if not title:
                title = clean(sh.text_frame.text)
    return {
        "title": title, "n_title": n_title, "n_subtitle": n_subtitle,
        "n_subhead": n_subhead, "other_ph": other_ph,
        "standalone_text": standalone_text, "groups": groups,
        "pictures": pictures, "ole": ole,
        "layout": slide.slide_layout.name,
    }


def descriptor(a):
    g, p, t = len(a["groups"]), a["pictures"], a["standalone_text"]
    gt = sum(a["groups"])
    if g and p:
        return f"{g}-item icon grid ({p} icons, caption per icon)"
    if g:
        return f"{g}-group diagram ({gt} text elements)"
    if p == 1 and t <= 2:
        return "text + 1 image layout"
    if p and t:
        return f"{t}-item grid with {p} icon(s)"
    if p:
        return f"{p}-image layout"
    if t >= 1:
        return f"{t}-item text layout"
    return "simple text layout"


def structure_str(a):
    parts = []
    if a["n_title"]:
        parts.append("title(idx0)")
    if a["n_subtitle"]:
        parts.append("subtitle(idx1)")
    if a["n_subhead"]:
        parts.append("subheadline(idx13)")
    for idx in a["other_ph"]:
        parts.append(f"content(idx{idx})")
    if a["standalone_text"]:
        parts.append(f"{a['standalone_text']} standalone text box(es)")
    if a["groups"]:
        inner = ", ".join(f"group of {n} text(s)" for n in a["groups"])
        parts.append(f"{len(a['groups'])} illustration group(s) ({inner})")
    if a["pictures"]:
        parts.append(f"{a['pictures']} picture(s)")
    if a["ole"]:
        parts.append(f"{a['ole']} OLE object (think-cell data, not clonable — skipped automatically)")
    parts.append(f"built on layout `{a['layout']}`")
    return "; ".join(parts)


def fillable(a):
    return a["standalone_text"] + sum(a["groups"]) + a["pictures"]


def section_map(prs):
    root = prs.part._element
    ns = {"p14": "http://schemas.microsoft.com/office/powerpoint/2010/main"}
    id_sec = {}
    order = []
    for sec in root.findall(".//p14:section", ns):
        for x in sec.findall(".//p14:sldId", ns):
            id_sec[x.get("id")] = sec.get("name")
        order.append(sec.get("name"))
    sldIdLst = list(prs.slides._sldIdLst)
    per_slide = []
    for s in sldIdLst:
        per_slide.append(id_sec.get(s.get("id"), "(unsectioned)"))
    return per_slide, order


def main():
    prs = Presentation(TEMPLATE)
    per_slide_sec, sec_order = section_map(prs)
    n = len(prs.slides)

    rows = []
    for i, slide in enumerate(prs.slides, 1):
        a = analyze(slide)
        sec = per_slide_sec[i - 1]
        desc = descriptor(a)
        is_master = sec == "Masterslides"
        simple = desc == "simple text layout" or (len(a["groups"]) == 0 and a["pictures"] == 0)
        title = a["title"]
        if is_placeholderish(title):
            label = f"{sec} — {desc}"
        else:
            label = f"{title} ({desc})"

        if is_master:
            use = (f"Base-layout demo (Masterslides) — clonable like any slide via "
                   f"`clone_slide: {i}`, but for a blank starting point you usually want "
                   f"the underlying layout `{a['layout']}` directly.")
        elif simple:
            use = (f"Simple content slide under '{sec}' — clone via `clone_slide: {i}`, "
                   f"or use its layout `{a['layout']}` directly for the same placeholder "
                   f"structure without the extra decoration.")
        else:
            use = (f"Hand-designed {desc} under '{sec}' — clone via `clone_slide: {i}` to "
                   f"reuse this exact composition; supply text_replacements for its "
                   f"{fillable(a)} fillable element(s) beyond the title/subheadline.")

        rows.append(f"| **Slide {i}** — {label} | {use} | {structure_str(a)} |")

    # section range summary
    ranges = []
    cur = None
    start = None
    for i, sec in enumerate(per_slide_sec, 1):
        if sec != cur:
            if cur is not None:
                ranges.append((cur, start, i - 1))
            cur, start = sec, i
    ranges.append((cur, start, n))

    header = f"""---
name: reference_library_index
description: Complete slide-by-slide index of all {n} example slides bundled in AFG_template.pptx, organized by section, with what each one is for and its exact clonable structure.
---

# Reference library index — all {n} slides in AFG_template.pptx

This is the full index referenced from `SKILL.md` and the system prompt.
**Check this file before building any slide** — not after something looks
repetitive, not as an afterthought. If a slide's content matches a row
here, clone that reference slide (`"clone_slide": <#>`) rather than
building from a plain layout, a procedural diagram, or bullets written
from scratch. Those remain the correct choice only when nothing in this
list is a genuine match for the content — they are the fallback, not the
default.

Every row was generated directly from the actual template file (title,
structure, and underlying layout all read from the real slide, not
hand-written) — if the template is ever updated, regenerate this file
rather than hand-editing it.

**Before cloning, run `catalog --template AFG_template.pptx --slide <#>`
first** to see the exact shape order `text_replacements` needs to
match — this table tells you WHICH slide to look at, `catalog` tells you
exactly HOW to fill it. If the count still ends up wrong anyway, the
build script's own warning includes the full shape-by-shape breakdown
(each box's current text and whether a replacement was supplied) directly
in its output, so a mismatch is fixable from that warning alone.

**Every one of these {n} slides is equally clonable — there is no
"verified" subset and no code-level restriction on any slide number,
including the Masterslides (base-layout demos) at the front.** Reach for
whichever reference slide actually fits the content, section by section;
don't fall back to the same one or two slides for everything, and don't
avoid a slide just because it isn't on some prior "known-good" list.
`clone_slide` edits the real slide's content in place (its
title/subheadline through placeholders, every decorative shape's text
through `text_replacements`) — full editing of the actual slide, not an
approximation.

**Real photography inside a reference slide stays exactly as designed.**
`clone_slide` clones every shape on the slide, including any actual photo
already embedded in it — that photo is part of the slide's designed
composition, not a placeholder to strip out or swap by default. Only
replace an embedded photo if the user explicitly asks for a different
image on that slide; otherwise leave it untouched and only supply
`text_replacements` for the text content.

## Sections at a glance ({len(ranges)} sections, {n} slides)

| Section | Slide range |
|---|---|
"""
    for name, s, e in ranges:
        rng = f"{s}" if s == e else f"{s}-{e}"
        header += f"| {name} | {rng} |\n"

    header += "\n| Slide (clone_slide: #) | Use for | Real placeholder structure |\n|---|---|---|\n"

    with open(OUT, "w") as f:
        f.write(header + "\n".join(rows) + "\n")
    print(f"Wrote {OUT}: {n} slide rows, {len(ranges)} sections")
    for name, s, e in ranges:
        print(f"  {name}: {s}-{e}")


if __name__ == "__main__":
    main()
