#!/usr/bin/env python3
"""
afg_deck_builder.py

Fixed, deterministic python-pptx logic for building and amending decks from
AFG_template.pptx. This script is meant to be called by the Copilot Studio
agent via Code Interpreter with a structured JSON plan as input -- the agent
should never write its own ad hoc placeholder-filling code. See SKILL.md for
the calling contract.

Usage (from Code Interpreter / CLI):

    python afg_deck_builder.py build   --template AFG_template.pptx --plan plan.json --output deck.pptx
    python afg_deck_builder.py amend   --input deck.pptx --plan amend_plan.json --output deck.pptx
    python afg_deck_builder.py inspect --input deck.pptx

Plan JSON schema (build):
{
  "slides": [
    {
      "layout": "Content Slide_2 columns_dark",
      "placeholders": {
        "0":  {"type": "text",  "value": "Why Power Apps Matters"},
        "13": {"type": "text",  "value": "Strategic value drivers"},
        "14": {"type": "bullets", "value": ["Faster delivery", "Lower cost"]},
        "15": {"type": "bullets", "value": ["Higher productivity"]}
      }
    },
    {
      "layout": "Split Content image_dark_1",
      "placeholders": {
        "0":  {"type": "text", "value": "Power Apps + Power Automate"},
        "13": {"type": "text", "value": "The workflow engine for daily efficiency"},
        "14": {"type": "diagram", "diagram_type": "process_steps", "items": [
          {"number": 1, "title": "Power Apps", "description": "Captures data via forms"},
          {"number": 2, "title": "Power Automate", "description": "Triggers approvals automatically"}
        ]},
        "15": {"type": "image", "value": "images/hero.jpg"}
      }
    },
    ...
  ]
}

Plan JSON schema (amend) -- list of targeted edits against an existing file:
{
  "edits": [
    {"slide_index": 3, "placeholders": {"14": {"type": "bullets", "value": [...]}}}
  ],
  "delete_slides": [7],
  "append_slides": [ { "layout": "...", "placeholders": {...} } ]
}

Every placeholder idx present in a slide's "placeholders" dict is filled.
Every placeholder idx that exists on the layout but is NOT present in the
dict is deleted from the slide (the mandatory cleanup pass). This makes
"dump everything into idx13" and "leave real content placeholders empty"
structurally impossible, because the caller must address each idx it wants
to keep -- anything unaddressed disappears rather than sitting there empty.

Both build and amend also run post-save composition checks (warnings, not
hard errors -- see SKILL.md "Layout and placeholder reference"):
  - check_duplicate_content  -- repeated verbatim text across slides
  - check_chart_coverage     -- deck has zero native charts
  - check_sparse_captions    -- 4-column icon captions under 4 words
                                 (bare labels like "Faster Delivery" that
                                 leave the rest of the slide looking empty)
  - check_layout_rhythm      -- 4+ consecutive slides using the same
                                 layout family with no image/chart slide
                                 breaking up the run
"""

import sys
import json
import copy
import re
import argparse
from lxml import etree
from pptx.oxml.ns import qn
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.dml.color import RGBColor

# AFG brand palette, pulled directly from AFG_template.pptx's theme1.xml
AFG_NAVY = RGBColor(0x0B, 0x1F, 0x41)
AFG_BLUE = RGBColor(0x00, 0xAE, 0xEF)
AFG_GREY = RGBColor(0xB3, 0xB3, 0xB3)
AFG_WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# Chart series colors and text color, chosen per slide theme so charts read
# clearly against both the dark-navy and light-white template variants.
CHART_THEME = {
    "dark":  {"series": [AFG_BLUE, AFG_WHITE, AFG_GREY], "text": AFG_WHITE},
    "light": {"series": [AFG_NAVY, AFG_BLUE, AFG_GREY],  "text": AFG_NAVY},
}

# Diagram (native-shape composition) colors, tinted off the same brand
# palette above so card grids / step chains sit comfortably on both slide
# themes without introducing new off-brand colors.
DIAGRAM_THEME = {
    "dark": {
        "card_fill": RGBColor(0x14, 0x2E, 0x5C),   # lighter navy tint, reads as a card against the darker slide bg
        "card_text": AFG_WHITE,
        "badge_fill": AFG_BLUE,
        "badge_text": AFG_WHITE,
        "connector": AFG_BLUE,
    },
    "light": {
        "card_fill": RGBColor(0xEA, 0xF6, 0xFD),   # pale blue tint against the white slide bg
        "card_text": AFG_NAVY,
        "badge_fill": AFG_NAVY,
        "badge_text": AFG_WHITE,
        "connector": AFG_BLUE,
    },
}

TEXT_TYPES = {"text", "bullets"}
IMAGE_TYPES = {"image"}
CHART_TYPES = {"chart"}
DIAGRAM_TYPES = {"diagram"}

CHART_KIND_MAP = {
    "bar": "COLUMN_CLUSTERED",
    "bar_horizontal": "BAR_CLUSTERED",
    "line": "LINE_MARKERS",
    "pie": "PIE",
}

DIAGRAM_KIND_MAP = {"process_steps", "card_grid", "accent_list", "stat_row"}
MIN_DIAGRAM_ITEMS = 2
MAX_DIAGRAM_ITEMS = 6

# Soft length guidance (warnings only, never blocks a build)
TITLE_WORD_LIMIT = 8
SUBHEAD_WORD_LIMIT = 15
BULLET_WORD_LIMIT = 20

# Rough glyph-width/line-height estimates for the TITLE placeholder (idx0)
# specifically -- its font is much larger/bolder than body or subheading
# text, so the same chars-per-line estimate used elsewhere badly
# undercounts wrapping. Calibrated against the real template: a title on
# a Split Content layout (title box sized for exactly one line, zero
# slack before the subheading starts right below it) visibly wrapped to 2
# lines and crowded the subheading -- these constants are chosen so the
# collision check catches that exact case.
TITLE_CHAR_WIDTH_PT = 15.0
TITLE_LINE_HEIGHT_PT = 38.0

# Same idea for the SUBHEADING placeholder (idx1/idx13): its font is
# smaller than the title's but still much bigger than body text, so a
# shared body-text estimate badly undercounts wrapping here too.
# Calibrated the same way as the title constant, against a real observed
# case: a 38-character subheading wrapped to 2 lines at ~15.9pt/char
# (line 1 broke after 27 characters within a ~430pt-wide box). The
# original constant here (7.5) was roughly half what's actually needed
# and essentially never fired.
SUBHEAD_CHAR_WIDTH_PT = 15.5
SUBHEAD_LINE_HEIGHT_PT = 30.0

# The template's actual default font sizes, confirmed directly from the
# theme XML (title placeholder's titleStyle defRPr sz="2800"; the
# subheading placeholder's own layout-level override also sz="2800") --
# both title and subheading default to 28pt. TITLE_CHAR_WIDTH_PT /
# SUBHEAD_CHAR_WIDTH_PT above were calibrated AT this size, so shrinking
# the font to fit one line scales proportionally from here.
TITLE_DEFAULT_PT = 28.0
SUBHEAD_DEFAULT_PT = 28.0
# Never shrink below this -- a title/subheading this much smaller than
# the template's design starts looking like a mistake in its own right;
# past this floor, leave it to wrap (and rely on check_title_subhead_
# collision / check_subhead_body_collision to flag it) rather than
# shrinking further.
TITLE_MIN_SHRINK_PT = 20.0

# Minimum words for a caption sitting next to an icon/picture on a compact
# multi-column layout. Below this, a caption reads as a bare label ("Faster
# Delivery") rather than a phrase explaining what it means or why it
# matters -- the single biggest cause of icon-grid slides that look sparse
# and leave the rest of the slide feeling empty (Section 2A of the system
# prompt). This is a warning, not a hard error, since a genuinely short
# proper-noun caption is occasionally correct.
MIN_CAPTION_WORDS = 4

# Content-placeholder idx values that sit paired with an icon/pic idx on
# "Content Slide_4 columns" layouts (pic idx -> its paired text idx).
FOUR_COLUMN_CAPTION_IDX = {14: 15, 16: 17, 18: 19, 20: 21}

# Layout "family" used to detect repetitive rhythm: strip the _dark/_light
# suffix so "Content Slide_4 columns_dark" and "...4 columns_light" count as
# the same shape. Consecutive runs of the same family beyond this length
# trigger a rhythm warning (Section 2 rotation rule / Section 2A). Set to 2
# (not 3) because in practice a *third* consecutive use of the same plain
# columns layout is already the point it starts reading as templated -- real
# decks (and Gamma's output) rarely repeat the identical slide shape twice in
# a row, let alone three times.
MAX_CONSECUTIVE_SAME_FAMILY = 2

# A "columns" layout counts as sparse when every content placeholder on it
# holds this many bullet paragraphs or fewer AND nothing else (no diagram,
# no chart, no image) fills the rest of the slide -- exactly the "3 bare
# columns of 3 short bullets, bottom half empty" pattern.
SPARSE_COLUMN_MAX_BULLETS = 5


def _layout_family(layout_name):
    name = layout_name
    for suffix in ("_dark", "_light"):
        if name.endswith(suffix):
            name = name[: -len(suffix)]
    return name


class PlanError(Exception):
    """Raised when the caller's plan doesn't match the template's reality."""
    pass


def _word_count(s):
    return len(s.split())


def get_layout_by_name(prs, name):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    available = ", ".join(l.name for l in prs.slide_layouts)
    raise PlanError(
        f"Layout '{name}' does not exist in this template. "
        f"Available layouts: {available}"
    )


def _placeholder_map(slide):
    """idx (int) -> placeholder shape, for every placeholder currently on the slide."""
    return {ph.placeholder_format.idx: ph for ph in slide.placeholders}


def _is_picture_ph(ph):
    # python-pptx assigns the PicturePlaceholder class for both PICTURE (18)
    # and BITMAP/clipArt (9) placeholder types -- checking the class instead
    # of matching the type enum's string name catches both correctly (the
    # dark "Title Page with image" layout uses BITMAP, not PICTURE, and a
    # substring match on "PICTURE" would incorrectly miss it).
    from pptx.shapes.placeholder import PicturePlaceholder
    return isinstance(ph, PicturePlaceholder)


def _is_picture_idx_on_layout(slide, idx):
    """
    Like _is_picture_ph, but checks the LAYOUT's declared placeholder type
    for this idx rather than the slide's current shape. Needed because
    insert_picture() replaces a PicturePlaceholder shape with a plain
    Picture shape (same idx, same visual slot) -- so _is_picture_ph on the
    slide's own shape returns False once an image has actually been
    inserted, even though the slot is still, structurally, the image slot.
    Any check that re-inspects an already-built/reloaded deck (rather than
    checking during the initial fill pass) needs this version.

    Also note: LAYOUT-level placeholders are all wrapped in python-pptx's
    generic LayoutPlaceholder class regardless of type -- isinstance
    against PicturePlaceholder (which only applies to SLIDE-level
    placeholders) never matches here, so this checks the type enum value
    directly instead.
    """
    from pptx.enum.shapes import PP_PLACEHOLDER
    for ph in slide.slide_layout.placeholders:
        if ph.placeholder_format.idx == idx:
            return ph.placeholder_format.type in (PP_PLACEHOLDER.PICTURE, PP_PLACEHOLDER.BITMAP)
    return False


BULLET_FONT_SIZE = Pt(20)
BULLET_SPACE_BEFORE_FIRST = Pt(22)   # extra gap between the subheading above and the first bullet
BULLET_SPACE_BEFORE_OTHERS = Pt(10)  # gap between individual bullet lines
BULLET_LINE_SPACING = 1.15


# Bullet marker glyphs the plan can request via "bullet_style" -- default
# stays "dot" (plain round bullet) unless the plan specifies otherwise.
# Use "arrow" for a sequence of related actions/capabilities (mirrors a
# style Gamma's own output uses), "check" for completed/included items,
# "dash" for a plainer, quieter list. This is glyph choice only -- size,
# color, and spacing stay governed by the existing bullet formatting
# rules below, so switching styles never turns into a manual formatting
# override.
BULLET_GLYPHS = {"dot": "\u2022", "arrow": "\u2192", "check": "\u2713", "dash": "\u2013"}
_BULLET_GLYPH_CHARS = tuple(BULLET_GLYPHS.values()) + ("•", "-", "*")


def _fit_single_line_font_size(text, width, default_pt, char_width_pt, min_pt):
    """
    Returns the font size (points) needed to keep `text` on one line within
    `width`, scaled down from `default_pt` if necessary, or None if it
    already fits at the default size. Character width is assumed to scale
    linearly with font size, calibrated from `char_width_pt` at
    `default_pt` (see TITLE_CHAR_WIDTH_PT / SUBHEAD_CHAR_WIDTH_PT). Never
    returns a value below `min_pt` -- past that floor, shrinking further
    starts looking like a mistake in its own right, so the caller should
    fall back to a warning instead of shrinking more.
    """
    if not width or not text:
        return None
    width_pt = width / Pt(1)
    cpl_at_default = width_pt / char_width_pt
    if len(text) <= cpl_at_default:
        return None  # fits already, no override needed
    required_pt = default_pt * (width_pt / (char_width_pt * len(text)))
    return max(required_pt, min_pt)


def _set_text(ph, value, warnings, slide_num, idx, bullet_style="dot"):
    """Fill a text-type placeholder. value may be a string (one line) or a list (bullets)."""
    tf = ph.text_frame
    tf.clear()  # leaves one empty paragraph
    if isinstance(value, list):
        lines = value
    else:
        lines = [value]

    is_bulleted = len(lines) > 1
    glyph = BULLET_GLYPHS.get(bullet_style, BULLET_GLYPHS["dot"])
    if bullet_style not in BULLET_GLYPHS:
        warnings.append(
            f"Slide {slide_num} idx {idx}: unknown bullet_style '{bullet_style}' -- "
            f"valid options are {sorted(BULLET_GLYPHS)}. Using 'dot'."
        )

    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        text = str(line)
        if is_bulleted:
            # Explicit bullet glyph -- more reliable across renderers than
            # editing buChar XML, and guarantees a visible marker instead of
            # bullets silently depending on inherited/undefined list formatting.
            if not text.lstrip().startswith(_BULLET_GLYPH_CHARS):
                text = f"{glyph}  {text}"
            p.space_before = BULLET_SPACE_BEFORE_FIRST if i == 0 else BULLET_SPACE_BEFORE_OTHERS
            p.line_spacing = BULLET_LINE_SPACING
        run = p.add_run()
        run.text = text
        if is_bulleted:
            run.font.size = BULLET_FONT_SIZE
        wc = _word_count(str(line))
        if len(lines) == 1 and wc > SUBHEAD_WORD_LIMIT and idx not in (14, 15, 16, 17, 19, 21):
            warnings.append(
                f"Slide {slide_num} idx {idx}: single-line text is {wc} words "
                f"(guideline: under {SUBHEAD_WORD_LIMIT})."
            )
        elif len(lines) > 1 and wc > BULLET_WORD_LIMIT:
            warnings.append(
                f"Slide {slide_num} idx {idx}: bullet '{line[:40]}...' is {wc} words "
                f"(guideline: under {BULLET_WORD_LIMIT})."
            )
        # Auto-shrink title (idx0) and subheading (idx1/13) to fit on one
        # line instead of just warning about wrap risk -- a title/
        # subheading box on this template is sized for exactly one line
        # with zero slack before whatever sits below it, so letting it
        # wrap silently produces a real, visible overlap (this happened in
        # practice, repeatedly, even after adding a warning for it -- a
        # warning only helps if it's read and acted on; this fixes it
        # unconditionally at render time instead). Only applies to
        # single-line values -- bulleted content is unaffected.
        if len(lines) == 1 and ph.width:
            if idx == 0:
                default_pt, char_width_pt = TITLE_DEFAULT_PT, TITLE_CHAR_WIDTH_PT
            elif idx in (1, 13):
                default_pt, char_width_pt = SUBHEAD_DEFAULT_PT, SUBHEAD_CHAR_WIDTH_PT
            else:
                default_pt = None
            if default_pt is not None:
                fitted_pt = _fit_single_line_font_size(
                    text, ph.width, default_pt, char_width_pt, TITLE_MIN_SHRINK_PT
                )
                if fitted_pt is not None:
                    run.font.size = Pt(fitted_pt)
                    # Only worth re-checking wrap risk when the floor was
                    # actually hit -- otherwise fitted_pt was derived
                    # exactly to make this text fit, and re-deriving the
                    # same formula here just re-hits floating-point
                    # rounding right at the boundary (len(text) landing on
                    # 30.00000001 instead of exactly 30.0), producing a
                    # false "still wraps" warning for a case that already
                    # fits cleanly.
                    if fitted_pt <= TITLE_MIN_SHRINK_PT:
                        effective_char_width = char_width_pt * (fitted_pt / default_pt)
                        if len(text) > (ph.width / Pt(1)) / effective_char_width:
                            label = "title" if idx == 0 else "subheadline"
                            warnings.append(
                                f"Slide {slide_num} idx {idx}: {label} '{text[:40]}' is "
                                f"still likely to wrap even after auto-shrinking to the "
                                f"minimum readable size -- shorten the text, it's too "
                                f"long for this placeholder's width."
                            )

    if is_bulleted and idx not in (0, 1, 10, 11, 12, 13):
        # Vertically center a short bullet list within its placeholder box
        # instead of leaving it pinned to the top -- a box sized to hold
        # many bullets but filled with only 2-4 short ones otherwise leaves
        # a large dead zone below the text (observed in practice). If the
        # list is actually long enough to fill the box, centering makes no
        # visible difference; if it's short, this is a meaningful improvement.
        # Title/subheading placeholders are excluded -- they should always
        # stay anchored where the template designed them.
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE


# Placeholder area (EMU^2) above which a slot counts as a "hero" image
# region (title banners, chapter/full-bleed/split-image panels) rather than
# a compact multi-column icon badge. Measured from the real template: the
# compact "Content Slide_4 columns" icon slot is ~4.8e12; every hero slot
# in the template is 4e13 or larger. Kept well below that gap so no
# legitimate compact slot is ever mistakenly blocked.
HERO_SLOT_AREA_THRESHOLD = 1.5e13


def _is_bundled_icon_path(image_path):
    normalized = image_path.replace("\\", "/")
    return "icons/" in normalized or normalized.startswith("icons/")


def _set_image(ph, image_path, warnings, slide_num, idx):
    if not _is_picture_ph(ph):
        raise PlanError(
            f"Slide {slide_num} idx {idx}: plan supplies an image but this placeholder "
            f"is not a picture placeholder. Never write an image into a text placeholder."
        )
    if _is_bundled_icon_path(image_path) and ph.width and ph.height:
        area = ph.width * ph.height
        if area > HERO_SLOT_AREA_THRESHOLD:
            raise PlanError(
                f"Slide {slide_num} idx {idx}: this placeholder is a large hero-style "
                f"image slot (title banner, chapter/full-bleed/split panel), not a "
                f"compact icon badge slot. A small bundled icon stretched into a slot "
                f"this size crops/blows up badly. Use a real photo or illustration for "
                f"this placeholder, or choose a different layout without a large image "
                f"slot instead -- do not put a bundled icons/ file here."
            )
    try:
        pic = ph.insert_picture(image_path)  # python-pptx crops to fill, preserving aspect ratio
    except Exception as e:
        raise PlanError(f"Slide {slide_num} idx {idx}: failed to insert image '{image_path}': {e}")

    # The layout's placeholder-prompt definition ("Click to insert
    # picture") carries its own solid fill (a visible gray/accent box
    # shown before an image is inserted). insert_picture() replaces the
    # shape with a real <p:pic> that has no fill of its own, but some
    # renderers still inherit that fill from the layout for anywhere the
    # inserted image doesn't fully opaque-cover -- which was invisible
    # with the skill's old fully-opaque icon PNGs, but became a visible
    # gray box behind every transparent icon (dark/light theme icons,
    # any transparent-background image) once those were introduced.
    # Explicitly disabling fill on the inserted picture prevents this
    # regardless of what the image's own transparency looks like.
    spPr = pic._element.spPr
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill"):
        for el in spPr.findall(qn(tag)):
            spPr.remove(el)
    etree.SubElement(spPr, qn("a:noFill"))


def _delete_placeholder(ph):
    sp = ph._element
    sp.getparent().remove(sp)


# ---------------------------------------------------------------------------
# Reference-visual cloning: copy an existing illustration (a shape or group)
# from one of the template's 102 reference slides, verbatim, into a new
# output slide, then only swap its text -- rather than redrawing an
# approximation of it with the procedural diagram engine below. This is the
# PRIORITY path whenever a matching illustration already exists in the
# template; the procedural diagrams (_draw_card_grid etc.) remain available
# for content shapes that don't have a close match in the reference library.
# ---------------------------------------------------------------------------

_REFERENCE_TEMPLATE_CACHE = {}


def _load_reference_template(template_path):
    """
    A second, untouched Presentation object opened purely for reading
    reference slides to clone from -- kept separate from the working
    presentation (which has its own slides cleared by
    _clear_existing_slides) so both can be used at once. Cached per path
    since a single build/amend run may clone from several reference slides.
    """
    if template_path not in _REFERENCE_TEMPLATE_CACHE:
        _REFERENCE_TEMPLATE_CACHE[template_path] = Presentation(template_path)
    return _REFERENCE_TEMPLATE_CACHE[template_path]


def _find_shape_by_id(slide, shape_id):
    for shape in slide.shapes:
        if shape.shape_id == shape_id:
            return shape
    return None


def _collect_text_shapes(shape, out):
    """Depth-first, document-order list of every text-bearing leaf shape
    inside a shape (or its own text frame if it's not a group) -- this
    fixed order is what a plan's text_replacements list is matched against,
    so cloning the same source slide always maps replacement #1 to the same
    physical box every time."""
    if shape.shape_type == 6:  # GROUP
        for sub in shape.shapes:
            _collect_text_shapes(sub, out)
    elif shape.has_text_frame and shape.text_frame.text.strip():
        out.append(shape)


def _rewrite_image_relationships(shape_element, source_part, target_part, media_cache):
    """
    Any picture nested inside the cloned shape/group references its image
    via an r:embed relationship ID that's only valid within the SOURCE
    slide part. After deep-copying the XML into the target slide, every
    such reference must be re-pointed at a new relationship added to the
    TARGET part (copying the underlying image bytes across once per unique
    image, cached in media_cache so a group with the same image reused
    twice doesn't duplicate the media part).
    """
    r_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    for blip in shape_element.iter(qn("a:blip")):
        old_rid = blip.get(f"{r_ns}embed")
        if not old_rid:
            continue
        cache_key = (id(source_part), old_rid)
        if cache_key not in media_cache:
            image_part = source_part.related_part(old_rid)
            new_rid = target_part.relate_to(image_part, image_part.reltype)
            media_cache[cache_key] = new_rid
        blip.set(f"{r_ns}embed", media_cache[cache_key])


def clone_reference_visual(target_slide, template_path, source_slide_index,
                            source_shape_id, placeholder_idx, left=None, top=None,
                            text_replacements=None, warnings=None):
    """
    Clones one shape (usually a GROUP containing a full illustration --
    donut segments, a timeline, a stat cascade, etc.) from a reference
    slide in the original template onto target_slide, verbatim -- same
    geometry, same colors, same effects -- then optionally repositions it
    and swaps its text.

    source_slide_index: 1-based slide number in the ORIGINAL template
      (i.e. as counted in AFG_template.pptx itself, before any slides are
      cleared) -- use the `catalog` command to find this and the matching
      source_shape_id for a given illustration.
    source_shape_id: the specific shape's .shape_id on that reference
      slide (a group id, usually) -- `catalog` lists these.
    left/top: optional EMU position to move the cloned shape to (e.g. to
      center it in a different placeholder's region than it originally
      occupied). If omitted, the shape keeps its original position, which
      is correct when cloning onto a similarly-proportioned layout.
    text_replacements: ordered list of strings. Matched, in document
      order, against every text-bearing leaf shape found inside the
      cloned shape (see _collect_text_shapes) -- replacement #1 goes into
      the first text box encountered, #2 into the second, and so on.
      Fewer replacements than text boxes leaves the remaining ones with
      their original template text (usually "Lorem ipsum..." placeholder
      copy) -- always supply a replacement for every text box, or the
      output will visibly contain leftover placeholder text.
    """
    warnings = warnings if warnings is not None else []
    ref_prs = _load_reference_template(template_path)
    try:
        ref_slide = ref_prs.slides[source_slide_index - 1]
    except IndexError:
        raise PlanError(
            f"clone_reference_visual: source_slide_index {source_slide_index} "
            f"does not exist in the template (it has {len(ref_prs.slides)} slides)."
        )
    source_shape = _find_shape_by_id(ref_slide, source_shape_id)
    if source_shape is None:
        raise PlanError(
            f"clone_reference_visual: shape id {source_shape_id} not found on "
            f"reference slide {source_slide_index}. Run the 'catalog' command "
            f"on this slide to list valid shape ids."
        )
    for shape in ref_slide.shapes:
        if shape.shape_type == 7 and shape.shape_id == source_shape_id:
            raise PlanError(
                f"clone_reference_visual: shape id {source_shape_id} on slide "
                f"{source_slide_index} is an embedded OLE object (think-cell "
                f"data), not a plain shape/group -- these can't be cloned this "
                f"way. Pick the shape id of the visual group instead (the "
                f"think-cell object is usually a separate top-level shape, not "
                f"part of the illustration's own group)."
            )

    new_element = copy.deepcopy(source_shape._element)
    media_cache = {}
    _rewrite_image_relationships(new_element, ref_slide.part, target_slide.part, media_cache)
    target_slide.shapes._spTree.append(new_element)

    # python-pptx doesn't wrap a raw appended element back into a Shape
    # object automatically -- re-find it by identity to get position/size
    # and text-shape access through the normal API.
    cloned = None
    for shape in target_slide.shapes:
        if shape._element is new_element:
            cloned = shape
            break

    if left is not None:
        cloned.left = left
    if top is not None:
        cloned.top = top

    # Tag it the same way charts/diagrams are tagged (see _tag_name), so
    # the empty-slide check and amend's _remove_generated_shapes cleanup
    # both recognize/find this as this idx's content.
    cloned.name = f"AFG_GEN_IDX{placeholder_idx}_CLONEDVISUAL_slide{source_slide_index}shape{source_shape_id}"

    if text_replacements is None:
        text_shapes = []
        _collect_text_shapes(cloned, text_shapes)
        if text_shapes:
            preview = "; ".join(
                repr(_clean_control_chars(s.text_frame.text.strip())[:25]) for s in text_shapes[:5]
            )
            more = f" (+{len(text_shapes) - 5} more)" if len(text_shapes) > 5 else ""
            warnings.append(
                f"clone_reference_visual (slide {source_slide_index}, shape "
                f"{source_shape_id}): no text_replacements given at all, but "
                f"this shape has {len(text_shapes)} text box(es) that still "
                f"contain the template's original placeholder text: "
                f"{preview}{more}. Run `catalog --slide {source_slide_index}` "
                f"and supply a text_replacements list, or this content ships "
                f"as unedited template placeholder text."
            )
    if text_replacements:
        text_shapes = []
        _collect_text_shapes(cloned, text_shapes)
        if len(text_replacements) != len(text_shapes):
            breakdown = []
            for idx, shape in enumerate(text_shapes):
                original = _clean_control_chars(shape.text_frame.text.strip())[:30]
                if idx < len(text_replacements):
                    status = f"-> will become {str(text_replacements[idx])[:30]!r}"
                else:
                    status = "-> NO replacement given, keeps template placeholder text"
                breakdown.append(f"    [{idx}] currently {original!r} {status}")
            direction = "fewer" if len(text_replacements) < len(text_shapes) else "more"
            warnings.append(
                f"clone_reference_visual (slide {source_slide_index}, shape "
                f"{source_shape_id}): {len(text_replacements)} replacement(s) "
                f"given but {len(text_shapes)} text box(es) actually found "
                f"({direction} replacements than boxes) -- exact shape-by-shape "
                f"breakdown, in the order text_replacements is matched "
                f"against:\n" + "\n".join(breakdown) +
                f"\nFix the text_replacements list to have exactly "
                f"{len(text_shapes)} entries in this order and rebuild."
            )
        for shape, new_text in zip(text_shapes, text_replacements):
            tf = shape.text_frame
            # Preserve the first run's formatting (font/size/color/bold),
            # replace only its text, and drop any extra runs/paragraphs so
            # multi-run placeholder text doesn't leave stale fragments.
            first_para = tf.paragraphs[0]
            if first_para.runs:
                first_para.runs[0].text = str(new_text)
                for extra_run in first_para.runs[1:]:
                    extra_run._r.getparent().remove(extra_run._r)
            else:
                first_para.add_run().text = str(new_text)
            for extra_para in tf.paragraphs[1:]:
                extra_para._p.getparent().remove(extra_para._p)
            # These text boxes are fixed-size, sized in the original design
            # for its own short placeholder copy ("Lorem ipsum dolor sit
            # amet") -- real replacement text is very often a different
            # length, and without this it silently overflows the box and
            # visually overlaps whatever neighboring shape sits below/beside
            # it (this happened in testing: longer sentences overlapped the
            # adjacent segment's number label). Shrinking to fit keeps text
            # inside its own box, which is more reliable than hoping every
            # replacement happens to be the same length as the placeholder.
            tf.word_wrap = True
            try:
                tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            except Exception:
                pass  # some shape/text-frame combinations don't support autofit; non-fatal

    return cloned


def clone_reference_slide(prs, template_path, source_slide_index,
                           placeholders=None, text_replacements=None,
                           slide_num=None, warnings=None):
    """
    Clones an ENTIRE reference slide (not just one shape/group) from the
    original template as a new slide in the working deck. Every
    non-placeholder, non-OLE shape on the source slide is copied verbatim
    -- position, formatting, images -- carrying over the whole
    hand-designed composition, rather than extracting a single
    illustration into a placeholder on a different layout (see
    clone_reference_visual for that narrower, targeted use). This is for
    genuinely using one of the reference library's 102 slides AS a slide,
    directly -- editing its own content in place -- not lifting one piece
    out of it.

    The new slide uses the SAME underlying layout the reference slide
    itself is built on, so its title/subheadline placeholders still exist
    and are filled normally through `placeholders` (identical to a
    regular slide spec) -- only the decorative/illustration shapes are
    cloned. OLE objects (think-cell data) can't be cloned and are skipped
    with a warning.

    text_replacements: ordered list of strings, matched in document order
    against every text-bearing leaf shape found across ALL cloned shapes
    (not counting the title/subheadline, which `placeholders` already
    covers) -- same matching rule as clone_reference_visual.
    """
    warnings = warnings if warnings is not None else []
    ref_prs = _load_reference_template(template_path)
    try:
        ref_slide = ref_prs.slides[source_slide_index - 1]
    except IndexError:
        raise PlanError(
            f"clone_reference_slide: source_slide_index {source_slide_index} "
            f"does not exist in the template (it has {len(ref_prs.slides)} slides)."
        )

    layout_name = ref_slide.slide_layout.name
    layout = get_layout_by_name(prs, layout_name)
    new_slide = prs.slides.add_slide(layout)

    if placeholders:
        fill_slide(new_slide, placeholders, slide_num or len(prs.slides), warnings, template_path=template_path)

    media_cache = {}
    all_text_shapes = []
    ole_skipped = 0
    for shape in ref_slide.shapes:
        if shape.is_placeholder:
            continue  # title/subheadline are handled above via `placeholders`
        if shape.shape_type == 7:  # embedded OLE object (think-cell data)
            ole_skipped += 1
            continue
        new_element = copy.deepcopy(shape._element)
        _rewrite_image_relationships(new_element, ref_slide.part, new_slide.part, media_cache)
        new_slide.shapes._spTree.append(new_element)
        cloned = None
        for s in new_slide.shapes:
            if s._element is new_element:
                cloned = s
                break
        if cloned is not None:
            cloned.name = f"AFG_GEN_WHOLESLIDE_{source_slide_index}_{shape.shape_id}"
            _collect_text_shapes(cloned, all_text_shapes)

    if ole_skipped:
        warnings.append(
            f"clone_reference_slide (slide {source_slide_index}): skipped "
            f"{ole_skipped} embedded OLE object(s) (think-cell data) that "
            f"can't be cloned -- if the illustration depends on that object "
            f"for its visible content, the result may be missing that piece; "
            f"verify visually before sharing the file."
        )

    if text_replacements is None and all_text_shapes:
        preview = "; ".join(
            repr(_clean_control_chars(s.text_frame.text.strip())[:25]) for s in all_text_shapes[:5]
        )
        more = f" (+{len(all_text_shapes) - 5} more)" if len(all_text_shapes) > 5 else ""
        warnings.append(
            f"clone_reference_slide (slide {source_slide_index}): no "
            f"text_replacements given at all, but this slide has "
            f"{len(all_text_shapes)} text box(es) that still contain the "
            f"template's original placeholder text: {preview}{more}. Run "
            f"`catalog --slide {source_slide_index}` and supply a "
            f"text_replacements list, or this content ships as unedited "
            f"template placeholder text."
        )
    if text_replacements is not None:
        if len(text_replacements) != len(all_text_shapes):
            breakdown = []
            for idx, shape in enumerate(all_text_shapes):
                original = _clean_control_chars(shape.text_frame.text.strip())[:30]
                if idx < len(text_replacements):
                    status = f"-> will become {str(text_replacements[idx])[:30]!r}"
                else:
                    status = "-> NO replacement given, keeps template placeholder text"
                breakdown.append(f"    [{idx}] currently {original!r} {status}")
            direction = "fewer" if len(text_replacements) < len(all_text_shapes) else "more"
            warnings.append(
                f"clone_reference_slide (slide {source_slide_index}): "
                f"{len(text_replacements)} replacement(s) given but "
                f"{len(all_text_shapes)} text box(es) actually found on the "
                f"source slide ({direction} replacements than boxes) -- exact "
                f"shape-by-shape breakdown, in the order text_replacements is "
                f"matched against:\n" + "\n".join(breakdown) +
                f"\nFix the text_replacements list to have exactly "
                f"{len(all_text_shapes)} entries in this order and rebuild -- "
                f"don't leave a mismatch and assume it's close enough."
            )
        for shape, new_text in zip(all_text_shapes, text_replacements):
            tf = shape.text_frame
            first_para = tf.paragraphs[0]
            if first_para.runs:
                first_para.runs[0].text = str(new_text)
                for extra_run in first_para.runs[1:]:
                    extra_run._r.getparent().remove(extra_run._r)
            else:
                first_para.add_run().text = str(new_text)
            for extra_para in tf.paragraphs[1:]:
                extra_para._p.getparent().remove(extra_para._p)
            tf.word_wrap = True
            try:
                tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            except Exception:
                pass  # some shape/text-frame combinations don't support autofit; non-fatal

    return new_slide


def _clean_control_chars(text):
    """
    Strips PowerPoint's internal control characters that are invisible in
    normal editors but invalid in plain text/markdown -- specifically
    0x0B (vertical tab, used internally as a soft line break within a
    single paragraph, e.g. many reference-template titles are literally
    "Please insert<0x0B>Chaptername") and 0x0C (form feed). This happened
    in practice: pulling slide titles directly from the template into a
    generated markdown table embedded these bytes verbatim, and the
    resulting .md file was rejected by an upload validator that checks
    for control characters. Any code that extracts text from a slide for
    use in documentation (not just this function) should route it
    through this first.
    """
    if text is None:
        return text
    return re.sub(r"[\x0b\x0c]", " ", text)


def catalog_reference_slide(template_path, slide_index):
    """
    Lists every top-level shape on a reference slide, with its shape_id,
    type, and a text preview -- and for groups, the ordered list of
    text-bearing sub-shapes exactly as clone_reference_visual will match
    them against a text_replacements list. Use this to find the right
    source_slide_index/source_shape_id before calling clone, and to know
    how many replacement strings to supply and in what order.
    """
    prs = _load_reference_template(template_path)
    try:
        slide = prs.slides[slide_index - 1]
    except IndexError:
        raise PlanError(f"catalog: slide_index {slide_index} does not exist.")
    lines = [f"Slide {slide_index} (layout: {slide.slide_layout.name}):"]
    for shape in slide.shapes:
        kind = "OLE (not clonable)" if shape.shape_type == 7 else str(shape.shape_type)
        preview = ""
        if shape.has_text_frame and shape.text_frame.text.strip():
            preview = _clean_control_chars(shape.text_frame.text.strip()).replace("\n", " | ")[:40]
        lines.append(f"  shape_id={shape.shape_id} [{kind}] {shape.name} {repr(preview)}")
        if shape.shape_type == 6:
            text_shapes = []
            _collect_text_shapes(shape, text_shapes)
            for i, ts in enumerate(text_shapes):
                preview_t = _clean_control_chars(ts.text_frame.text.strip())[:40]
                lines.append(f"      text_replacements[{i}] <- {repr(preview_t)}")
    return "\n".join(lines)


def _detect_theme(slide):
    """Infer dark/light theme from the slide's layout name (e.g. '..._dark')."""
    name = slide.slide_layout.name.lower()
    return "dark" if "_dark" in name else "light"


def _tag_name(idx, kind, n=None):
    """
    Deterministic shape name for anything a chart/diagram generates in place
    of a placeholder (graphic frames, cards, badges, text boxes). These
    shapes are NOT placeholders once created, so a later amend can't find
    them via the normal placeholder map -- they need their own lookup key
    so a re-amend of the same idx can find and remove the old ones instead
    of drawing new shapes on top of stale ones (see _remove_generated_shapes).
    """
    base = f"AFG_GEN_IDX{idx}_{kind}"
    return base if n is None else f"{base}_{n}"


def _remove_generated_shapes(slide, idx):
    """
    Removes every shape previously tagged with _tag_name for this idx --
    call this before re-filling an idx that might previously have held a
    chart or diagram, since those replace the placeholder with freestanding
    shapes that a simple ph_map-based restore/refill can't see or clean up
    on its own. Safe to call even if nothing was ever generated for this idx.
    """
    prefix = f"AFG_GEN_IDX{idx}_"
    for shape in list(slide.shapes):
        if shape.name.startswith(prefix):
            shape._element.getparent().remove(shape._element)


def _set_chart(slide, ph, chart_spec, warnings, slide_num, idx):
    """
    Replaces the given placeholder with a native, editable pptx chart at the
    exact same position/size the placeholder occupied -- so charts stay
    aligned to the template's grid instead of floating at arbitrary
    coordinates. Works for any placeholder type (text or picture), since the
    chart fully replaces the shape rather than filling into it.

    Colors, title/axis text color, data labels, and axis titles are all set
    explicitly rather than left at python-pptx/PowerPoint defaults, since the
    default Office chart theme (blue/orange/grey) has poor contrast against
    this template's navy dark-mode slides and doesn't match the brand
    palette on light slides either.
    """
    kind = chart_spec.get("chart_type", "bar")
    if kind not in CHART_KIND_MAP:
        raise PlanError(
            f"Slide {slide_num} idx {idx}: unknown chart_type '{kind}'. "
            f"Valid options: {sorted(CHART_KIND_MAP.keys())}"
        )
    xl_type = getattr(XL_CHART_TYPE, CHART_KIND_MAP[kind])

    categories = chart_spec.get("categories")
    series = chart_spec.get("series")
    if not categories or not series:
        raise PlanError(
            f"Slide {slide_num} idx {idx}: chart requires both 'categories' "
            f"(list) and 'series' (dict of name -> list of numbers)."
        )

    chart_data = CategoryChartData()
    chart_data.categories = categories
    for series_name, values in series.items():
        if len(values) != len(categories):
            raise PlanError(
                f"Slide {slide_num} idx {idx}: series '{series_name}' has "
                f"{len(values)} values but there are {len(categories)} categories."
            )
        chart_data.add_series(series_name, values)

    left, top, width, height = ph.left, ph.top, ph.width, ph.height
    if left is None or top is None:
        raise PlanError(
            f"Slide {slide_num} idx {idx}: could not resolve this placeholder's "
            f"position to place the chart. Try a different idx on this layout."
        )

    # Reserve a thin strip at the bottom for a source citation, if given,
    # so the chart itself doesn't overlap the caption.
    source = chart_spec.get("source")
    if source:
        caption_h = Pt(16)
        height = height - caption_h

    _delete_placeholder(ph)
    graphic_frame = slide.shapes.add_chart(xl_type, left, top, width, height, chart_data)
    graphic_frame.name = _tag_name(idx, "CHART")
    chart = graphic_frame.chart

    theme = _detect_theme(slide)
    palette = CHART_THEME[theme]
    text_color = palette["text"]

    # Series colors from the AFG palette, cycled if there are more series
    # than palette entries.
    plot = chart.plots[0]
    for i, series_obj in enumerate(plot.series):
        color = palette["series"][i % len(palette["series"])]
        if kind == "pie":
            # Color each individual slice (pie has one series, many points)
            for j, point in enumerate(series_obj.points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = palette["series"][j % len(palette["series"])]
        else:
            series_obj.format.fill.solid()
            series_obj.format.fill.fore_color.rgb = color
            if kind == "line":
                series_obj.format.line.color.rgb = color
                series_obj.format.line.width = Pt(2.5)

    chart.has_legend = len(series) > 1 and kind != "pie"
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.color.rgb = text_color
        chart.legend.font.size = Pt(11)
    if kind == "pie":
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.RIGHT
        chart.legend.include_in_layout = False
        chart.legend.font.color.rgb = text_color
        chart.legend.font.size = Pt(11)

    title = chart_spec.get("title")
    if title:
        chart.has_title = True
        chart.chart_title.text_frame.text = title
        for p in chart.chart_title.text_frame.paragraphs:
            for r in p.runs:
                r.font.color.rgb = text_color
                r.font.size = Pt(14)
                r.font.bold = True
    else:
        chart.has_title = False

    # Data labels -- the actual numbers on bars/points/slices, not just axes.
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.font.size = Pt(10)
    data_labels.font.color.rgb = text_color
    if kind == "pie":
        data_labels.number_format_is_linked = False
        data_labels.number_format = '0'
        data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    else:
        data_labels.position = XL_LABEL_POSITION.OUTSIDE_END

    # Axis text color and optional axis titles (category/value axes don't
    # exist on pie charts).
    if kind != "pie":
        cat_axis = chart.category_axis
        val_axis = chart.value_axis
        cat_axis.tick_labels.font.color.rgb = text_color
        cat_axis.tick_labels.font.size = Pt(11)
        val_axis.tick_labels.font.color.rgb = text_color
        val_axis.tick_labels.font.size = Pt(11)
        cat_axis.format.line.color.rgb = text_color
        val_axis.format.line.color.rgb = text_color

        x_title = chart_spec.get("x_axis_title")
        if x_title:
            cat_axis.has_title = True
            cat_axis.axis_title.text_frame.text = x_title
            for p in cat_axis.axis_title.text_frame.paragraphs:
                for r in p.runs:
                    r.font.color.rgb = text_color
                    r.font.size = Pt(10)

        y_title = chart_spec.get("y_axis_title")
        if y_title:
            val_axis.has_title = True
            val_axis.axis_title.text_frame.text = y_title
            for p in val_axis.axis_title.text_frame.paragraphs:
                for r in p.runs:
                    r.font.color.rgb = text_color
                    r.font.size = Pt(10)

    # Small source citation strip beneath the chart -- required whenever the
    # figures came from a web source rather than the user directly.
    if source:
        cap_box = slide.shapes.add_textbox(left, top + height, width, Pt(16))
        cap_box.name = _tag_name(idx, "CHART_CAPTION")
        cap_tf = cap_box.text_frame
        cap_tf.word_wrap = True
        cap_p = cap_tf.paragraphs[0]
        cap_run = cap_p.add_run()
        cap_run.text = f"Source: {source}"
        cap_run.font.size = Pt(8)
        cap_run.font.italic = True
        cap_run.font.color.rgb = text_color




def _set_diagram(slide, ph, spec, warnings, slide_num, idx):
    """
    Replaces the given placeholder with a native-shape composition -- built
    from real PowerPoint autoshapes (rounded rectangles, ovals, thin accent
    bars) at the exact position/size the placeholder occupied. This is the
    tool for turning a list of parallel points into an arranged illustration
    instead of a plain bulleted list (Section 2A / Section 3 of the system
    prompt): every shape is native and editable, not an image, so it stays
    on-brand and on-grid.

    Four diagram types, deliberately different visual languages so a deck
    doesn't read as "the same box, over and over" -- see SKILL.md's
    "Diagrams" section for the full variety-selection guidance:
      - "process_steps": vertical chain of numbered circular badges.
      - "card_grid": one or more filled rounded-rect cards.
      - "accent_list": vertical list with a thin colored accent bar per
        item -- no boxes at all. A lighter-weight alternative to card_grid
        for capability/feature lists that don't need a boxed container.
      - "stat_row": a horizontal row of big typographic numbers with a
        thin accent underline and a label -- for illustrative metrics
        that don't warrant a full chart. Also box-free.

    process_steps/card_grid/accent_list items: {"title": str,
    "description": str?, "bullets": [str]?, "number": int? (process_steps
    only)}. Give either "description" or "bullets" per item, never both.
    stat_row items: {"value": str, "label": str} (e.g. {"value": "70%",
    "label": "Faster delivery"}).

    process_steps needs 2-6 items; the others need 1-6.
    """
    diagram_type = spec.get("diagram_type")
    if diagram_type not in DIAGRAM_KIND_MAP:
        raise PlanError(
            f"Slide {slide_num} idx {idx}: unknown diagram_type '{diagram_type}'. "
            f"Valid options: {sorted(DIAGRAM_KIND_MAP)}"
        )
    items = spec.get("items")
    if not items:
        raise PlanError(f"Slide {slide_num} idx {idx}: diagram requires a non-empty 'items' list.")
    # process_steps AND accent_list both distribute items evenly across the
    # FULL placeholder height by construction -- with only 1 item, that
    # means "full height / 1", stretching a single badge/accent-bar into
    # an oddly tall shape with the text floating in the middle of a mostly
    # empty column (this happened in practice: a single-item accent_list
    # used for the "one category per already-narrow layout column"
    # pattern). card_grid is the right tool for that pattern instead --
    # it sizes to content rather than stretching, so 1 item is fine there.
    min_items = MIN_DIAGRAM_ITEMS if diagram_type in ("process_steps", "accent_list") else 1
    if len(items) < min_items:
        raise PlanError(
            f"Slide {slide_num} idx {idx}: '{diagram_type}' diagram needs at least "
            f"{min_items} item(s)."
        )
    if len(items) > MAX_DIAGRAM_ITEMS:
        warnings.append(
            f"Slide {slide_num} idx {idx}: diagram has {len(items)} items, more than "
            f"{MAX_DIAGRAM_ITEMS} -- consider splitting across two slides, each item "
            f"will be cramped at this count."
        )

    left, top, width, height = ph.left, ph.top, ph.width, ph.height
    if left is None or top is None:
        raise PlanError(
            f"Slide {slide_num} idx {idx}: could not resolve this placeholder's "
            f"position to place the diagram. Try a different idx on this layout."
        )
    # Same top buffer bullets get (BULLET_SPACE_BEFORE_FIRST) -- without it,
    # a subheadline that wraps to two lines on a narrower column (e.g. one
    # side of a Split Content layout) visually collides with the first
    # card/badge, since the placeholder's raw top edge assumes a one-line
    # subheadline sitting above it.
    top = top + BULLET_SPACE_BEFORE_FIRST
    height = height - BULLET_SPACE_BEFORE_FIRST

    theme = _detect_theme(slide)
    palette = DIAGRAM_THEME[theme]

    if diagram_type == "stat_row":
        for item in items:
            value = item.get("value")
            if not value or not item.get("label"):
                warnings.append(
                    f"Slide {slide_num} idx {idx}: stat_row item missing 'value' or "
                    f"'label' -- both are required (e.g. value='70%', label='Faster delivery')."
                )
                continue
            value = str(value)
            # A "value" longer than ~6 characters with no digit/%/x at all
            # is almost always a misuse -- a KPI NAME ("Revenue", "Margin")
            # put in the number slot instead of a short metric. This
            # happened in practice and, combined with 3+ items, produced
            # visibly overlapping text between columns. The fix isn't a
            # bigger box -- it's the right tool: a KPI name + category
            # belongs in card_grid (title + description), not stat_row.
            if len(value) > 6 and not any(c.isdigit() for c in value) and "%" not in value:
                warnings.append(
                    f"Slide {slide_num} idx {idx}: stat_row 'value' is '{value}' -- "
                    f"this looks like a KPI name/label, not a short metric. 'value' "
                    f"should be a number or short metric (e.g. '70%', '3x', '15'); a "
                    f"named KPI card (title + description) belongs in 'card_grid' "
                    f"instead, not 'stat_row'."
                )
    else:
        for item in items:
            desc = item.get("description") or ""
            bullets = item.get("bullets") or []
            wc = _word_count(desc)
            if wc > BULLET_WORD_LIMIT:
                warnings.append(
                    f"Slide {slide_num} idx {idx}: diagram item '{item.get('title','')[:30]}' "
                    f"description is {wc} words (guideline: under {BULLET_WORD_LIMIT})."
                )
            for b in bullets:
                bwc = _word_count(str(b))
                if bwc > BULLET_WORD_LIMIT:
                    warnings.append(
                        f"Slide {slide_num} idx {idx}: diagram item '{item.get('title','')[:30]}' "
                        f"bullet '{str(b)[:30]}' is {bwc} words (guideline: under {BULLET_WORD_LIMIT})."
                    )
            if not desc and not bullets:
                warnings.append(
                    f"Slide {slide_num} idx {idx}: diagram item '{item.get('title','')[:30]}' "
                    f"has no 'description' or 'bullets' -- a title-only item reproduces the "
                    f"bare-label sparseness problem in shape form. Add at least one of them."
                )
            if desc and bullets:
                warnings.append(
                    f"Slide {slide_num} idx {idx}: diagram item '{item.get('title','')[:30]}' "
                    f"has both 'description' and 'bullets' -- only 'bullets' will be rendered. "
                    f"Use one or the other."
                )

    _delete_placeholder(ph)
    if diagram_type == "process_steps":
        _draw_process_steps(slide, left, top, width, height, items, palette, idx)
    elif diagram_type == "accent_list":
        _draw_accent_list(slide, left, top, width, height, items, palette, idx)
    elif diagram_type == "stat_row":
        _draw_stat_row(slide, left, top, width, height, items, palette, idx)
    else:
        _draw_card_grid(slide, left, top, width, height, items, palette, idx)


def _est_lines(text, chars_per_line, cap):
    """Rough ceil-division line-wrap estimate, capped so one absurdly long
    string can't blow the height budget out to something silly."""
    if not text:
        return 0
    return min(max(1, -(-len(str(text)) // chars_per_line)), cap)


def _draw_card_grid(slide, left, top, width, height, items, palette, idx):
    """
    Horizontal row of equal-width rounded-rect cards. Each card holds a
    title plus EITHER a one-line "description" (a single sentence) OR a
    short "bullets" list (2-4 short items) -- use "bullets" for grouped
    list content (e.g. "Retail Use Cases" broken into 3 categories of 3
    items each), which reads far better as titled bullet cards than as
    three bare columns of plain text with nothing else on the slide.

    Cards are sized to their actual content -- title line-wrap included,
    not just the description/bullets -- so a title that wraps to 2 lines
    (very possible in a 3+ card row on a narrower column) still fits
    without pushing text past the card's bottom edge.

    Each card gets a thin colored accent bar across its top edge
    automatically (matching the reference library's fact-card treatment).
    Optionally give an item an `"icon"` field (a path into `icons/` or
    `images/`) to place a small icon beside its title -- use this the same
    way `accent_list`'s icon field works: only when it genuinely adds
    meaning, and consistently across every card in the same grid (all of
    them get an icon, or none do).

    Dynamic space fill: if the placeholder has meaningfully more height
    than the cards need at their base size, font sizes/padding scale up
    (capped) so the cards genuinely fill more of the available space
    instead of sitting as a small block with a large dead zone below --
    the row is also vertically centered in that case, rather than always
    hugging the top.
    """
    n = len(items)
    gap = Pt(14)
    card_w = (width - gap * (n - 1)) // n
    inner_w = card_w - Pt(24)  # minus left+right margins, for wrap estimates
    has_icons = any(item.get("icon") for item in items)
    icon_col_w = Pt(30) if has_icons else Pt(0)

    def compute_height(scale):
        # Distinct chars-per-line estimates per text style -- bold title
        # glyphs are wider on average than body/bullet glyphs, so a shared
        # estimate under-counts title wrapping specifically (the slide 8
        # bug: a 2-line title pushed the description past the card's
        # bottom edge). Larger scale -> fewer chars fit per line.
        title_cpl = max(int((inner_w - icon_col_w) / (Pt(8.2) * scale)), 6)
        body_cpl = max(int(inner_w / (Pt(6.2) * scale)), 8)
        bullet_cpl = max(int((inner_w - Pt(10)) / (Pt(6.0) * scale)), 6)

        def item_height(item):
            title_lines = _est_lines(item.get("title", ""), title_cpl, cap=2)
            h = title_lines * Pt(18 * scale)
            bullets = item.get("bullets")
            desc = item.get("description") or item.get("desc")
            if bullets:
                h += Pt(6 * scale)
                for b in bullets:
                    h += _est_lines(b, bullet_cpl, cap=2) * Pt(14 * scale) + Pt(5 * scale)
            elif desc:
                h += Pt(5 * scale) + _est_lines(desc, body_cpl, cap=4) * Pt(15 * scale)
            return h + Pt(24 * scale)

        return max(item_height(item) for item in items)

    natural_h = compute_height(1.0)
    # Grow fonts/padding to use spare vertical room, capped at 1.6x so text
    # never balloons past a sensible size -- this is the fix for cards that
    # were technically non-overlapping but left half the slide empty below
    # them: instead of a fixed small size, the card block scales toward the
    # space it's actually given.
    #
    # Scale is also capped by available WIDTH per card, not just height --
    # a card_grid paired with an image (so each card is narrow, sharing
    # roughly half the slide) has plenty of spare HEIGHT but very little
    # spare WIDTH; scaling fonts up purely based on height overflowed those
    # narrow cards (a word like "Microsoft" broke mid-word because the
    # scaled-up font no longer fit the card's width at all). The width cap
    # shrinks smoothly as cards get narrower, so a wide card_grid (e.g. one
    # alone on a full-width layout) can still scale up fully.
    card_w_pt = card_w / Pt(1)
    width_cap = max(1.0, min(1.6, card_w_pt / 140.0))
    scale = 1.0
    if height > natural_h * 1.15:
        scale = min(height / natural_h, 1.6, width_cap)
    card_h = min(height, compute_height(scale))

    # Only center vertically when we deliberately scaled up to fill space;
    # the compact (scale == 1) case keeps the original top-anchored
    # behavior, which is correct when a card_grid is paired with other
    # content below/beside it (e.g. an image) that already balances the slide.
    row_top = top + (height - card_h) // 2 if scale > 1.01 else top + Pt(4)

    title_size = min(Pt(14 * scale), Pt(22))
    desc_size = min(Pt(11 * scale), Pt(16))
    bullet_size = min(Pt(10.5 * scale), Pt(15))
    margin_scale = min(scale, 1.4)

    for i, item in enumerate(items):
        card_left = left + i * (card_w + gap)
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card_left, row_top, card_w, card_h)
        card.name = _tag_name(idx, "CARD", i)
        card.adjustments[0] = 0.08
        card.fill.solid()
        card.fill.fore_color.rgb = palette["card_fill"]
        card.line.fill.background()
        card.shadow.inherit = False

        # Thin colored accent bar across the top edge of the card -- the
        # single detail that reads as "designed" rather than "a plain box",
        # matching the reference library's fact-card treatment.
        accent_h = Pt(4)
        accent = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, card_left, row_top, card_w, accent_h
        )
        accent.name = _tag_name(idx, "ACCENTBAR", i)
        accent.fill.solid()
        accent.fill.fore_color.rgb = palette["badge_fill"]
        accent.line.fill.background()
        accent.shadow.inherit = False

        tf = card.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_top = Pt(12 * margin_scale) + accent_h
        tf.margin_bottom = Pt(10 * margin_scale)

        icon_path = item.get("icon")
        title_indent = Pt(0)
        if icon_path:
            icon_size = Pt(20 * margin_scale)
            icon_left = card_left + Pt(12 * margin_scale)
            icon_top = row_top + Pt(12 * margin_scale) + accent_h
            try:
                pic = slide.shapes.add_picture(icon_path, icon_left, icon_top, icon_size, icon_size)
                pic.name = _tag_name(idx, "CARDICON", i)
                title_indent = icon_size + Pt(8 * margin_scale)
            except Exception:
                title_indent = Pt(0)  # missing/bad icon path -- fall back to no indent rather than failing the build

        tf.margin_left = Pt(12 * margin_scale) + title_indent
        tf.margin_right = Pt(12 * margin_scale)

        p_title = tf.paragraphs[0]
        r_title = p_title.add_run()
        r_title.text = str(item.get("title", ""))
        r_title.font.bold = True
        r_title.font.size = title_size
        r_title.font.color.rgb = palette["card_text"]

        bullets = item.get("bullets")
        desc = item.get("description") or item.get("desc")
        if bullets:
            for b in bullets:
                p_b = tf.add_paragraph()
                p_b.space_before = Pt(5 * margin_scale)
                p_b.line_spacing = 1.05
                r_b = p_b.add_run()
                text = str(b)
                if not text.lstrip().startswith(("•", "-", "*")):
                    text = f"\u2022  {text}"
                r_b.text = text
                r_b.font.size = bullet_size
                r_b.font.color.rgb = palette["card_text"]
        elif desc:
            p_desc = tf.add_paragraph()
            p_desc.space_before = Pt(5 * margin_scale)
            p_desc.line_spacing = 1.1
            r_desc = p_desc.add_run()
            r_desc.text = str(desc)
            r_desc.font.size = desc_size
            r_desc.font.color.rgb = palette["card_text"]


def _draw_accent_list(slide, left, top, width, height, items, palette, idx):
    """
    Vertical list, no boxes: each item gets either a small bundled icon
    (if "icon" is given) or a thin colored accent bar on the left, then a
    bold title and a description/bullets to its right -- evenly
    distributed across the FULL available height (like process_steps), so
    this type always fills the space by construction rather than needing
    a scale-up heuristic. Use this as a lighter-weight alternative to
    card_grid/process_steps when you want visual variety on a slide near
    others that already used a boxed or numbered treatment, and use the
    per-item "icon" field when each point genuinely has a distinct,
    relevant icon (e.g. a specific capability or workflow step) -- don't
    add icons just to add them; the plain accent bar is the right default.
    """
    n = len(items)
    row_gap = Pt(14)
    row_h = (height - row_gap * (n - 1)) // n
    icon_size = min(row_h, Pt(32))
    bar_w = Pt(5)

    for i, item in enumerate(items):
        row_top = top + i * (row_h + row_gap)
        icon_path = item.get("icon")

        if icon_path:
            icon_top = row_top + (row_h - icon_size) // 2
            pic = slide.shapes.add_picture(icon_path, left, icon_top, icon_size, icon_size)
            pic.name = _tag_name(idx, "ACCENT_ICON", i)
            text_left = left + icon_size + Pt(16)
            text_width = width - icon_size - Pt(16)
        else:
            bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, row_top, bar_w, row_h)
            bar.name = _tag_name(idx, "ACCENTBAR", i)
            bar.adjustments[0] = 0.5
            bar.fill.solid()
            bar.fill.fore_color.rgb = palette["badge_fill"]
            bar.line.fill.background()
            bar.shadow.inherit = False
            text_left = left + bar_w + Pt(16)
            text_width = width - bar_w - Pt(16)

        text_box = slide.shapes.add_textbox(text_left, row_top, text_width, row_h)
        text_box.name = _tag_name(idx, "ACCENT_TEXT", i)
        ttf = text_box.text_frame
        ttf.word_wrap = True
        ttf.vertical_anchor = MSO_ANCHOR.MIDDLE

        p_title = ttf.paragraphs[0]
        r_title = p_title.add_run()
        r_title.text = str(item.get("title", ""))
        r_title.font.bold = True
        r_title.font.size = Pt(15)
        r_title.font.color.rgb = palette["card_text"]

        bullets = item.get("bullets")
        desc = item.get("description") or item.get("desc")
        if bullets:
            for b in bullets:
                p_b = ttf.add_paragraph()
                p_b.space_before = Pt(3)
                p_b.line_spacing = 1.05
                r_b = p_b.add_run()
                text = str(b)
                if not text.lstrip().startswith(("•", "-", "*")):
                    text = f"\u2022  {text}"
                r_b.text = text
                r_b.font.size = Pt(11)
                r_b.font.color.rgb = palette["card_text"]
        elif desc:
            p_desc = ttf.add_paragraph()
            p_desc.space_before = Pt(3)
            p_desc.line_spacing = 1.1
            r_desc = p_desc.add_run()
            r_desc.text = str(desc)
            r_desc.font.size = Pt(11.5)
            r_desc.font.color.rgb = palette["card_text"]


def _draw_stat_row(slide, left, top, width, height, items, palette, idx):
    """
    Horizontal row of big typographic numbers, no boxes: each item shows a
    large "value" (e.g. "70%"), a thin accent underline, and a "label"
    beneath. For illustrative metrics that don't need a full chart (see
    Charts section for when a real chart is required instead). Font size
    scales with the available height/width so a stat row in a tall/wide
    placeholder reads as a genuine typographic moment, not small numbers
    floating in a big empty area. All sizing math below works in plain
    point values (floats) and only converts to EMU via Pt() at the point
    of use, to avoid mixing units.
    """
    n = len(items)
    gap = Pt(20)
    col_w = (width - gap * (n - 1)) // n

    height_pt = height / Pt(1)
    col_w_pt = col_w / Pt(1)

    value_size_pt = min(40.0, height_pt / 3.6, col_w_pt / 3.2)
    value_size_pt = max(value_size_pt, 22.0)

    # Also cap by each item's actual value text length -- the height/width
    # caps above assume a short metric like "70%" or "3x". If "value" ends
    # up holding a longer word instead (a real misuse -- see the validation
    # warning in _set_diagram -- but the render still needs to not visibly
    # break), a fixed font size sized for "70%" overflows into the next
    # column. ~0.58 is an approximate average bold-character-width-to-font-size
    # ratio; dividing the column width by (chars * that ratio) gives the
    # largest font size that keeps the text inside its own column.
    for item in items:
        value_text = str(item.get("value", ""))
        if value_text:
            max_for_this_item = col_w_pt / (len(value_text) * 0.58)
            value_size_pt = min(value_size_pt, max_for_this_item)
    value_size_pt = max(value_size_pt, 14.0)  # absolute floor so it never becomes illegible

    label_size_pt = min(max(value_size_pt / 3, 11.0), 14.0)

    # Vertically center the whole value+bar+label block in the available
    # height when there's room to spare, rather than always hugging the
    # top and leaving a large dead zone below (the same fix applied to
    # card_grid's scale-up case).
    block_h_pt = value_size_pt + 14 + 4 + 10 + label_size_pt * 2.2
    top = top + max(Pt(0), (height - Pt(block_h_pt)) // 2)

    for i, item in enumerate(items):
        col_left = left + i * (col_w + gap)

        value_box = slide.shapes.add_textbox(col_left, top, col_w, Pt(value_size_pt) + Pt(10))
        value_box.name = _tag_name(idx, "STATVALUE", i)
        vtf = value_box.text_frame
        vtf.word_wrap = False
        vp = vtf.paragraphs[0]
        vp.alignment = PP_ALIGN.LEFT
        vr = vp.add_run()
        vr.text = str(item.get("value", ""))
        vr.font.bold = True
        vr.font.size = Pt(value_size_pt)
        vr.font.color.rgb = palette["badge_fill"]

        bar_top = top + Pt(value_size_pt) + Pt(14)
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col_left, bar_top, min(col_w, Pt(40)), Pt(4))
        bar.name = _tag_name(idx, "STATBAR", i)
        bar.adjustments[0] = 0.5
        bar.fill.solid()
        bar.fill.fore_color.rgb = palette["connector"]
        bar.line.fill.background()
        bar.shadow.inherit = False

        label_top = bar_top + Pt(10)
        label_box = slide.shapes.add_textbox(col_left, label_top, col_w, height - (label_top - top))
        label_box.name = _tag_name(idx, "STATLABEL", i)
        ltf = label_box.text_frame
        ltf.word_wrap = True
        lp = ltf.paragraphs[0]
        lr = lp.add_run()
        lr.text = str(item.get("label", ""))
        lr.font.size = Pt(label_size_pt)
        lr.font.color.rgb = palette["card_text"]


def _draw_process_steps(slide, left, top, width, height, items, palette, idx):
    """Vertical chain of numbered circular badges, each paired with title+description."""
    n = len(items)
    row_gap = Pt(12)
    row_h = (height - row_gap * (n - 1)) // n
    badge_size = min(row_h, Pt(40))
    text_left = left + badge_size + Pt(16)
    text_width = width - badge_size - Pt(16)

    for i, item in enumerate(items):
        row_top = top + i * (row_h + row_gap)
        badge_top = row_top + (row_h - badge_size) // 2

        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, badge_top, badge_size, badge_size)
        badge.name = _tag_name(idx, "BADGE", i)
        badge.fill.solid()
        badge.fill.fore_color.rgb = palette["badge_fill"]
        badge.line.fill.background()
        badge.shadow.inherit = False
        btf = badge.text_frame
        btf.word_wrap = False
        btf.vertical_anchor = MSO_ANCHOR.MIDDLE
        btf.margin_left = 0
        btf.margin_right = 0
        btf.margin_top = 0
        btf.margin_bottom = 0
        bp = btf.paragraphs[0]
        bp.alignment = PP_ALIGN.CENTER
        br = bp.add_run()
        br.text = str(item.get("number", i + 1))
        br.font.bold = True
        br.font.size = Pt(16)
        br.font.color.rgb = palette["badge_text"]

        text_box = slide.shapes.add_textbox(text_left, row_top, text_width, row_h)
        text_box.name = _tag_name(idx, "STEP_TEXT", i)
        ttf = text_box.text_frame
        ttf.word_wrap = True
        ttf.vertical_anchor = MSO_ANCHOR.MIDDLE

        tp = ttf.paragraphs[0]
        tr = tp.add_run()
        tr.text = str(item.get("title", ""))
        tr.font.bold = True
        tr.font.size = Pt(14)
        tr.font.color.rgb = palette["card_text"]

        desc = item.get("description") or item.get("desc")
        if desc:
            tp2 = ttf.add_paragraph()
            tp2.space_before = Pt(3)
            tp2.line_spacing = 1.1
            tr2 = tp2.add_run()
            tr2.text = str(desc)
            tr2.font.size = Pt(11)
            tr2.font.color.rgb = palette["card_text"]


def _center_four_column_icons(slide, slide_height):
    """
    "Content Slide_4 columns" positions its icon+caption placeholders at a
    fixed position defined by the layout, sized for the icon/caption
    content itself -- not stretched to fill the slide. When that content is
    short (the normal case -- a caption is a phrase, not a paragraph), the
    row sits pinned near the top with a large dead zone below it (the
    "slide 3 looks cramped up top" bug). This vertically centers the whole
    icon+caption row within the space between the subheading and the
    slide's bottom margin, matching how card_grid/process_steps diagrams
    already fill their own placeholder -- so this native icon layout gets
    the same treatment instead of being the one layout left out of it.
    Safe no-op if there isn't meaningful slack to redistribute.
    """
    ph_map = _placeholder_map(slide)
    icon_idxs = [14, 16, 18, 20]
    cap_idxs = [15, 17, 19, 21]
    present_icons = [i for i in icon_idxs if i in ph_map]
    present_caps = [i for i in cap_idxs if i in ph_map]
    if not present_icons or not present_caps:
        return  # already converted away from icons (e.g. amended into something else)

    tops = [ph_map[i].top for i in present_icons if ph_map[i].top is not None]
    # Use the caption's actual TEXT height, not its fixed placeholder box
    # height -- the box is sized generously for a longer caption than most
    # actually use, so measuring the box bottom understates how much real
    # slack there is and produces a smaller shift than the slide visually
    # has room for.
    cap_bottoms = []
    for i in present_caps:
        ph = ph_map[i]
        if ph.top is None or ph.width is None:
            continue
        text = ph.text_frame.text if ph.has_text_frame else ""
        cpl = max(int(ph.width / Pt(6.5)), 8)
        lines = _est_lines(text, cpl, cap=6) if text else 1
        text_h = lines * Pt(16) + Pt(8)
        box_h = ph.height if ph.height is not None else text_h
        cap_bottoms.append(ph.top + min(text_h, box_h))
    if not tops or not cap_bottoms:
        return

    block_top = min(tops)
    block_bottom = max(cap_bottoms)
    block_h = block_bottom - block_top
    if block_h <= 0:
        return

    sub_ph = ph_map.get(13) or ph_map.get(1)
    region_top = block_top
    if sub_ph is not None and sub_ph.top is not None and sub_ph.height is not None:
        region_top = sub_ph.top + sub_ph.height + BULLET_SPACE_BEFORE_FIRST

    region_bottom = slide_height - Emu(457200)  # ~0.5in bottom margin, clear of the footer
    available = region_bottom - region_top
    slack = available - block_h
    if slack < Pt(20):
        return  # negative, zero, or too small to bother with

    dy = slack // 2
    for i in present_icons + present_caps:
        ph = ph_map[i]
        # Snapshot ALL FOUR geometry values before mutating any of them.
        # Setting one field (e.g. .top) on a placeholder with no explicit
        # transform of its own creates a brand-new transform and defaults
        # every OTHER field to 0 -- so reading .top again AFTER setting
        # .left would return the already-reset 0, not the original value.
        # Every placeholder ended up with the identical top (== dy) the
        # first time this was written, because each one's "original top"
        # was read post-reset. Snapshotting first avoids that entirely.
        orig_left, orig_top, orig_width, orig_height = ph.left, ph.top, ph.width, ph.height
        if orig_top is None:
            continue
        ph.left = orig_left
        ph.top = orig_top + dy
        ph.width = orig_width
        ph.height = orig_height


_SINGLE_CARD_NAME_RE = re.compile(r"^AFG_GEN_IDX(\d+)_CARD_0$")


def _normalize_sibling_cards(slide):
    """
    Finds single-item card_grid results on the same slide (one card per
    column idx -- e.g. AFG_GEN_IDX14_CARD_0, AFG_GEN_IDX15_CARD_0,
    AFG_GEN_IDX16_CARD_0, the "one category per already-narrow column"
    pattern) and normalizes them to identical height and top position.

    Each single-item card_grid call sizes independently based on its own
    content -- so if one column's title/bullets happen to need slightly
    more or less wrapped text than its siblings, that card ends up a
    different height AND a different vertical position (since taller
    cards also get more/less centering offset), even though they're
    meant to read as parallel, uniform cards in a row. This happened in
    practice: 3 single-item cards on a "Content Slide_3 columns" slide
    rendered at visibly different sizes. Only matches the CARD_0 tag
    (single-item form) -- multi-item card_grid rows are already drawn
    together in one call and are already uniform by construction.
    """
    card_shapes = []
    for shape in slide.shapes:
        if _SINGLE_CARD_NAME_RE.match(getattr(shape, "name", "") or ""):
            if shape.top is not None and shape.height is not None:
                card_shapes.append(shape)
    if len(card_shapes) < 2:
        return
    common_top = min(s.top for s in card_shapes)
    common_height = max(s.height for s in card_shapes)
    for s in card_shapes:
        s.top = common_top
        s.height = common_height


def _add_slide_numbers(prs):
    """
    Draws an explicit slide-number text box in the bottom-right corner of
    every slide, colored to contrast with that slide's dark/light theme
    (white on dark, navy on light -- same palette _detect_theme already
    drives for chart/diagram coloring). The template's layouts define a
    slide-number PLACEHOLDER (idx11) with exactly this positioning and
    theme-aware color already built in, but python-pptx's add_slide()
    does not clone footer/slide-number/date placeholders onto a new slide
    the way it does for ordinary content placeholders -- they're simply
    never present on the slide to fill or preserve, which is why slide
    numbers never appeared even though the template looks like it
    supports them. Drawing a plain text box directly is the reliable fix.
    Runs once after all slides are finalized (build or amend, after any
    slide inserts/deletes/reorders) so numbering always reflects the
    final slide order. Idempotent across repeated amends -- any
    previously-drawn slide-number box is removed and redrawn fresh each
    time, rather than stacking duplicates.
    """
    box_w = Pt(50)
    box_h = Pt(22)
    margin = Pt(16)
    for i, slide in enumerate(prs.slides, start=1):
        for shape in list(slide.shapes):
            if getattr(shape, "name", "") == "AFG_SLIDE_NUMBER":
                shape._element.getparent().remove(shape._element)
        theme = _detect_theme(slide)
        color = AFG_WHITE if theme == "dark" else AFG_NAVY
        left = prs.slide_width - box_w - margin
        top = prs.slide_height - box_h - margin
        box = slide.shapes.add_textbox(left, top, box_w, box_h)
        box.name = "AFG_SLIDE_NUMBER"
        tf = box.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        run = p.add_run()
        run.text = str(i)
        run.font.size = Pt(12)
        run.font.color.rgb = color


def _ensure_closing_slide_is_last(prs):
    """
    Moves the last "Thank you and contact" slide (if any) to the true end
    of the slide order, if it isn't already there. python-pptx's
    add_slide() always appends to the very end of the slide list with no
    "insert at position" option -- so amend's append_slides silently
    stranded a closing Thank You slide in the middle of the deck when new
    content was added afterward (this happened in practice: 9 new slides
    landed after an existing Thank You slide instead of before it). This
    runs before slide numbering so numbers reflect the corrected order,
    and applies to both build and amend since the rule -- the deck always
    closes with Thank You, never mid-deck -- has no legitimate exception.
    """
    thank_you_positions = [
        i for i, s in enumerate(prs.slides)
        if _layout_family(s.slide_layout.name) == "Thank you and contact"
    ]
    if not thank_you_positions:
        return
    last_position = thank_you_positions[-1]
    if last_position == len(prs.slides) - 1:
        return  # already last, nothing to do
    xml_slides = prs.slides._sldIdLst
    slide_ids = list(xml_slides)
    moving = slide_ids[last_position]
    xml_slides.remove(moving)
    xml_slides.append(moving)


def _apply_layout_polish(prs):
    """
    Post-process pass run once after all slides are built/amended -- for
    fixes that need the slide in its final state and apply per-layout
    regardless of which specific plan filled it. Currently: moving a
    stranded Thank You slide back to the true end (see
    _ensure_closing_slide_is_last, run FIRST so slide numbering below
    reflects the corrected order), vertically centering "Content Slide_4
    columns" icon rows (see _center_four_column_icons), normalizing
    single-item card_grid siblings to a uniform size (see
    _normalize_sibling_cards), and drawing slide numbers (see
    _add_slide_numbers). Extend here for future per-layout polish rather
    than threading more logic through fill_slide's per-placeholder loop.
    """
    _ensure_closing_slide_is_last(prs)
    for slide in prs.slides:
        family = _layout_family(slide.slide_layout.name)
        if family == "Content Slide_4 columns":
            _center_four_column_icons(slide, prs.slide_height)
        _normalize_sibling_cards(slide)
    _add_slide_numbers(prs)


def fill_slide(slide, placeholders_plan, slide_num, warnings, template_path=None):
    """
    placeholders_plan: dict of {idx(str or int): {"type": "text"|"bullets"|"image"|"chart"|"diagram"|"cloned_visual", "value": ...}}
    Fills every idx present in the plan. Deletes every placeholder idx present on the
    slide but NOT present in the plan (mandatory cleanup pass -- Section 4.2 of the
    system prompt).
    template_path: required only if any placeholder uses "type": "cloned_visual"
    (it needs to re-open the original template to read the reference slide from).
    """
    ph_map = _placeholder_map(slide)
    plan_idxs = {int(k) for k in placeholders_plan.keys()}

    # 1. Validate every requested idx actually exists on this layout
    missing = plan_idxs - set(ph_map.keys())
    if missing:
        raise PlanError(
            f"Slide {slide_num}: plan references idx {sorted(missing)} which do not "
            f"exist on this layout. Available idx on this layout: {sorted(ph_map.keys())}"
        )

    # 2. Fill requested placeholders, type-checked
    for idx_str, spec in placeholders_plan.items():
        idx = int(idx_str)
        ph = ph_map[idx]
        kind = spec.get("type")
        value = spec.get("value")

        if kind in TEXT_TYPES:
            if _is_picture_ph(ph):
                raise PlanError(
                    f"Slide {slide_num} idx {idx}: plan supplies text for a picture "
                    f"placeholder. Never write text into a pic/clipArt placeholder."
                )
            _set_text(ph, value, warnings, slide_num, idx, spec.get("bullet_style", "dot"))
        elif kind in IMAGE_TYPES:
            _set_image(ph, value, warnings, slide_num, idx)
        elif kind in CHART_TYPES:
            _set_chart(slide, ph, spec, warnings, slide_num, idx)
        elif kind in DIAGRAM_TYPES:
            _set_diagram(slide, ph, spec, warnings, slide_num, idx)
        elif kind == "cloned_visual":
            if not template_path:
                raise PlanError(
                    f"Slide {slide_num} idx {idx}: 'cloned_visual' requires "
                    f"template_path to be passed through to fill_slide."
                )
            left, top, width, height = ph.left, ph.top, ph.width, ph.height
            _delete_placeholder(ph)
            cloned = clone_reference_visual(
                slide, template_path,
                source_slide_index=spec["source_slide"],
                source_shape_id=spec["source_shape_id"],
                placeholder_idx=idx,
                text_replacements=spec.get("text_replacements"),
                warnings=warnings,
            )
            if spec.get("fit_to_placeholder", True) and left is not None:
                # Scale the cloned group to fit the placeholder's box
                # (preserving its aspect ratio) and center it there, so it
                # lines up with this layout's grid instead of keeping
                # whatever size it was on its original reference slide.
                orig_w, orig_h = cloned.width, cloned.height
                if orig_w and orig_h:
                    scale = min(width / orig_w, height / orig_h)
                    cloned.width = int(orig_w * scale)
                    cloned.height = int(orig_h * scale)
                    cloned.left = left + (width - cloned.width) // 2
                    cloned.top = top + (height - cloned.height) // 2
        else:
            raise PlanError(f"Slide {slide_num} idx {idx}: unknown placeholder type '{kind}'.")

    # 3. Cleanup pass -- delete every placeholder NOT addressed by the plan
    unused_idxs = set(ph_map.keys()) - plan_idxs
    for idx in unused_idxs:
        _delete_placeholder(ph_map[idx])


def _clear_existing_slides(prs):
    """
    AFG_template.pptx ships with a large reference library of example
    slides (organized into named PowerPoint Sections purely as a browsing
    aid for picking the right layout for a given purpose -- see SKILL.md
    "Reference library" note). If we don't strip these before adding new
    content, they'd silently remain at the front of every generated deck.
    This removes all slides but leaves every layout/master untouched.

    It also removes the presentation's <p14:sectionLst> (PowerPoint
    "Sections" feature), if present. That list references specific slide
    IDs, and every one of those IDs belongs to a reference slide we just
    deleted -- leaving stale section entries behind would produce a file
    with dangling references, which risks a "repair needed" prompt when
    the output is opened in real PowerPoint. The section grouping is only
    meaningful in the source template as a reference aid; it has no
    purpose in a freshly generated output deck.
    """
    xml_slides = prs.slides._sldIdLst
    slide_ids = list(xml_slides)
    for sldId in slide_ids:
        xml_slides.remove(sldId)
        # also drop the now-orphaned slide part to keep the file clean
        rId = sldId.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        prs.part.drop_rel(rId)

    root = prs.part._element
    ns = {"p14": "http://schemas.microsoft.com/office/powerpoint/2010/main"}
    for section_lst in root.findall(".//p14:sectionLst", ns):
        ext = section_lst.getparent()  # <p:ext uri="...">
        ext_lst = ext.getparent()      # <p:extLst>
        ext_lst.remove(ext)


def build_deck(template_path, plan, output_path):
    prs = Presentation(template_path)
    _clear_existing_slides(prs)
    warnings = []

    for i, slide_spec in enumerate(plan["slides"], start=1):
        if "clone_slide" in slide_spec:
            clone_reference_slide(
                prs, template_path,
                source_slide_index=slide_spec["clone_slide"],
                placeholders=slide_spec.get("placeholders", {}),
                text_replacements=slide_spec.get("text_replacements"),
                slide_num=i, warnings=warnings,
            )
        else:
            layout = get_layout_by_name(prs, slide_spec["layout"])
            slide = prs.slides.add_slide(layout)
            fill_slide(slide, slide_spec.get("placeholders", {}), i, warnings, template_path=template_path)

    _remove_empty_slides(prs, warnings)
    _apply_layout_polish(prs)

    dup_warnings = check_duplicate_content(prs)
    warnings.extend(dup_warnings)
    warnings.extend(check_title_subhead_collision(prs))
    warnings.extend(check_subhead_body_collision(prs))
    warnings.extend(check_chart_coverage(prs))
    warnings.extend(check_chart_variety(prs))
    warnings.extend(check_no_reference_library_usage(prs))
    warnings.extend(check_all_charts_illustrative(prs))
    warnings.extend(check_sparse_captions(prs))
    warnings.extend(check_layout_rhythm(prs))
    warnings.extend(check_sparse_columns(prs))
    warnings.extend(check_diagram_variety(prs))
    warnings.extend(check_narrow_diagram_full_width_layout(prs))
    warnings.extend(check_empty_slide(prs))
    warnings.extend(check_chapter_before_thank_you(prs))
    warnings.extend(check_mixed_icon_usage(prs))
    warnings.extend(check_mixed_bullet_styles_on_slide(prs))
    warnings.extend(check_generic_contact_info(prs))

    _block_on_duplicate_body_content(prs)
    prs.save(output_path)
    return summarize(prs, output_path, warnings)


def _restore_placeholder_from_layout(slide, idx):
    """
    If a placeholder idx was previously deleted from a slide (e.g. by an
    earlier cleanup pass) but the slide's layout still defines it, clone it
    back from the layout so it can be filled again. Returns the restored
    placeholder, or None if the layout doesn't define this idx either.
    """
    layout = slide.slide_layout
    layout_ph = None
    for ph in layout.placeholders:
        if ph.placeholder_format.idx == idx:
            layout_ph = ph
            break
    if layout_ph is None:
        return None
    new_sp = copy.deepcopy(layout_ph._element)
    slide.shapes._spTree.append(new_sp)
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            return ph
    return None


def amend_deck(input_path, plan, output_path, reference_template_path=None):
    """
    reference_template_path: the ORIGINAL AFG_template.pptx (with its full
    102-slide reference library) -- required only if any edit or appended
    slide uses "type": "cloned_visual". This is deliberately NOT the same
    as input_path: input_path is the user's in-progress working deck (which
    only has that deck's own content slides, no reference library left in
    it, since it was already stripped by _clear_existing_slides when it was
    first built) -- cloning always needs to read from the untouched
    original template, never from the working deck being amended.
    """
    prs = Presentation(input_path)
    warnings = []

    for edit in plan.get("edits", []):
        slide_index = edit["slide_index"]  # 1-based, matches how slides are presented to the user
        try:
            slide = prs.slides[slide_index - 1]
        except IndexError:
            raise PlanError(f"amend: slide_index {slide_index} does not exist (deck has {len(prs.slides)} slides).")
        # Amend fills only the idx explicitly given; it does NOT delete other
        # placeholders on the slide that weren't mentioned in this edit, since
        # an amendment is a partial change, not a full rebuild of that slide.
        ph_map = _placeholder_map(slide)
        for idx_str, spec in edit.get("placeholders", {}).items():
            idx = int(idx_str)
            # A previous build/amend may have replaced this idx with a chart
            # or diagram -- those are freestanding shapes, not placeholders,
            # so ph_map can't see them. Clear any such leftover shapes for
            # this idx before restoring/refilling, or the new content ends
            # up drawn on top of the old (see _remove_generated_shapes).
            _remove_generated_shapes(slide, idx)
            if idx not in ph_map:
                restored = _restore_placeholder_from_layout(slide, idx)
                if restored is None:
                    raise PlanError(
                        f"amend slide {slide_index}: idx {idx} does not exist on this "
                        f"slide or its layout."
                    )
                ph_map[idx] = restored
                warnings.append(
                    f"Slide {slide_index} idx {idx}: this placeholder had been removed "
                    f"previously and was restored from the layout before filling."
                )
            ph = ph_map[idx]
            kind = spec.get("type")
            value = spec.get("value")
            if kind in TEXT_TYPES:
                if _is_picture_ph(ph):
                    raise PlanError(f"amend slide {slide_index} idx {idx}: text into a picture placeholder is not allowed.")
                _set_text(ph, value, warnings, slide_index, idx, spec.get("bullet_style", "dot"))
            elif kind in IMAGE_TYPES:
                _set_image(ph, value, warnings, slide_index, idx)
            elif kind in CHART_TYPES:
                _set_chart(slide, ph, spec, warnings, slide_index, idx)
            elif kind in DIAGRAM_TYPES:
                _set_diagram(slide, ph, spec, warnings, slide_index, idx)
            elif kind == "cloned_visual":
                if not reference_template_path:
                    raise PlanError(
                        f"amend slide {slide_index} idx {idx}: 'cloned_visual' requires "
                        f"reference_template_path (the original AFG_template.pptx) to be "
                        f"passed to amend_deck."
                    )
                left, top, width, height = ph.left, ph.top, ph.width, ph.height
                _delete_placeholder(ph)
                cloned = clone_reference_visual(
                    slide, reference_template_path,
                    source_slide_index=spec["source_slide"],
                    source_shape_id=spec["source_shape_id"],
                    placeholder_idx=idx,
                    text_replacements=spec.get("text_replacements"),
                    warnings=warnings,
                )
                if spec.get("fit_to_placeholder", True) and left is not None:
                    orig_w, orig_h = cloned.width, cloned.height
                    if orig_w and orig_h:
                        scale = min(width / orig_w, height / orig_h)
                        cloned.width = int(orig_w * scale)
                        cloned.height = int(orig_h * scale)
                        cloned.left = left + (width - cloned.width) // 2
                        cloned.top = top + (height - cloned.height) // 2
            else:
                raise PlanError(f"amend slide {slide_index} idx {idx}: unknown type '{kind}'.")
        for idx in edit.get("delete_placeholders", []):
            _remove_generated_shapes(slide, idx)
            if idx in ph_map:
                _delete_placeholder(ph_map[idx])

    # Deletions (by 1-based index, highest first so indices don't shift mid-loop)
    for slide_index in sorted(plan.get("delete_slides", []), reverse=True):
        xml_slides = prs.slides._sldIdLst
        slide_ids = list(xml_slides)
        xml_slides.remove(slide_ids[slide_index - 1])

    # Appends (new slides added at the end, same rules as build)
    for slide_spec in plan.get("append_slides", []):
        if "clone_slide" in slide_spec:
            clone_reference_slide(
                prs, reference_template_path,
                source_slide_index=slide_spec["clone_slide"],
                placeholders=slide_spec.get("placeholders", {}),
                text_replacements=slide_spec.get("text_replacements"),
                slide_num=len(prs.slides) + 1, warnings=warnings,
            )
        else:
            layout = get_layout_by_name(prs, slide_spec["layout"])
            slide = prs.slides.add_slide(layout)
            fill_slide(slide, slide_spec.get("placeholders", {}), len(prs.slides), warnings, template_path=reference_template_path)

    _remove_empty_slides(prs, warnings)
    _apply_layout_polish(prs)

    dup_warnings = check_duplicate_content(prs)
    warnings.extend(dup_warnings)
    warnings.extend(check_title_subhead_collision(prs))
    warnings.extend(check_subhead_body_collision(prs))
    warnings.extend(check_chart_coverage(prs))
    warnings.extend(check_chart_variety(prs))
    warnings.extend(check_no_reference_library_usage(prs))
    warnings.extend(check_all_charts_illustrative(prs))
    warnings.extend(check_sparse_captions(prs))
    warnings.extend(check_layout_rhythm(prs))
    warnings.extend(check_sparse_columns(prs))
    warnings.extend(check_diagram_variety(prs))
    warnings.extend(check_narrow_diagram_full_width_layout(prs))
    warnings.extend(check_empty_slide(prs))
    warnings.extend(check_chapter_before_thank_you(prs))
    warnings.extend(check_mixed_icon_usage(prs))
    warnings.extend(check_mixed_bullet_styles_on_slide(prs))
    warnings.extend(check_generic_contact_info(prs))

    _block_on_duplicate_body_content(prs)
    prs.save(output_path)
    return summarize(prs, output_path, warnings)


def _x_ranges_overlap(left1, width1, left2, width2):
    """
    True if two horizontal ranges overlap at all. Used to avoid treating a
    shape in a completely different column as a collision risk just
    because its Y position happens to be lower -- this was a real false
    positive: a layout's contact-info block sat in a separate column to
    the right of the subheading, with no horizontal overlap at all, but
    a Y-only check still flagged it as "content the subheading could
    overlap." When either shape's geometry is unknown, returns True
    (assume overlap) since a missed real collision is worse than an extra
    warning to double check.
    """
    if left1 is None or width1 is None or left2 is None or width2 is None:
        return True
    right1, right2 = left1 + width1, left2 + width2
    return left1 < right2 and left2 < right1


def _get_first_run_font_pt(ph, default_pt):
    """
    Returns the actual font size (points) of the first run in this
    placeholder if explicitly set, else `default_pt`. Used so the
    post-build collision checks reflect any auto-shrink already applied
    by _set_text (via _fit_single_line_font_size) instead of assuming
    every title/subheading is still at the template's default size.
    """
    if ph.has_text_frame:
        for para in ph.text_frame.paragraphs:
            for run in para.runs:
                if run.font.size is not None:
                    return run.font.size / Pt(1)
    return default_pt


def check_title_subhead_collision(prs):
    """
    Flags a title that's likely to wrap to 2+ lines at its placeholder's
    actual width AND, based on that estimated wrapped height, would
    crowd or overlap the subheading placeholder directly below it. This
    is a real bug observed in practice: a title box sized for exactly one
    line (common on Split Content layouts, where the title column is only
    half the slide's width) wrapped to 2 lines and visually collided with
    the subheading right below it. _set_text now auto-shrinks the font to
    prevent this at render time -- this check is a backstop for the rare
    case where even the minimum readable size doesn't fit (using the
    ACTUAL applied font size, not always the template default, so it
    doesn't re-flag something the shrink already fixed). Warning only --
    the estimate is a heuristic, not a pixel-perfect layout engine.
    """
    warnings = []
    for i, slide in enumerate(prs.slides, start=1):
        ph_map = _placeholder_map(slide)
        title_ph = ph_map.get(0)
        if title_ph is None or not title_ph.has_text_frame or not title_ph.width:
            continue
        title_text = title_ph.text_frame.text.strip()
        if not title_text or title_ph.top is None:
            continue

        actual_pt = _get_first_run_font_pt(title_ph, TITLE_DEFAULT_PT)
        scale = actual_pt / TITLE_DEFAULT_PT
        cpl = max(int(title_ph.width / Pt(TITLE_CHAR_WIDTH_PT * scale)), 6)
        est_lines = _est_lines(title_text, cpl, cap=4)
        if est_lines < 2:
            continue
        est_title_bottom = title_ph.top + est_lines * Pt(TITLE_LINE_HEIGHT_PT * scale)

        sub_ph = ph_map.get(13) or ph_map.get(1)
        if (
            sub_ph is not None
            and sub_ph.top is not None
            and _x_ranges_overlap(title_ph.left, title_ph.width, sub_ph.left, sub_ph.width)
        ):
            if est_title_bottom > sub_ph.top:
                warnings.append(
                    f"Slide {i}: title '{title_text[:40]}' is likely to wrap to "
                    f"{est_lines} lines at this placeholder's width, which will crowd "
                    f"or overlap the subheading directly below it. Shorten the title "
                    f"-- especially on narrower layouts (e.g. Split Content, where the "
                    f"title column is only about half the slide's width) -- or verify "
                    f"visually before sharing the file."
                )
        elif title_ph.height and est_title_bottom > title_ph.top + title_ph.height:
            warnings.append(
                f"Slide {i}: title '{title_text[:40]}' is likely to wrap to "
                f"{est_lines} lines and overflow its own text box at this width. "
                f"Shorten it."
            )
    return warnings


def check_subhead_body_collision(prs):
    """
    Same idea as check_title_subhead_collision, one level down: flags a
    subheading that's likely to wrap to 2+ lines at its placeholder's
    actual width in a way that would overlap whatever body content
    (bullets, a diagram, a chart) starts directly below it. This is a
    real bug observed in practice on a Split Content layout: the body
    placeholder's top position assumes a one-line subheading above it,
    so a wrapped 2-line subheading visually overlapped the first bullet.
    """
    warnings = []
    for i, slide in enumerate(prs.slides, start=1):
        ph_map = _placeholder_map(slide)
        sub_ph = ph_map.get(13) or ph_map.get(1)
        if sub_ph is None or not sub_ph.has_text_frame or not sub_ph.width or sub_ph.top is None:
            continue
        sub_text = sub_ph.text_frame.text.strip()
        if not sub_text:
            continue

        cpl = max(int(sub_ph.width / Pt(SUBHEAD_CHAR_WIDTH_PT * (_get_first_run_font_pt(sub_ph, SUBHEAD_DEFAULT_PT) / SUBHEAD_DEFAULT_PT))), 6)
        est_lines = _est_lines(sub_text, cpl, cap=3)
        if est_lines < 2:
            continue
        scale = _get_first_run_font_pt(sub_ph, SUBHEAD_DEFAULT_PT) / SUBHEAD_DEFAULT_PT
        est_sub_bottom = sub_ph.top + est_lines * Pt(SUBHEAD_LINE_HEIGHT_PT * scale)

        body_tops = []
        for idx, ph in ph_map.items():
            if idx in (0, 1, 10, 11, 12, 13):
                continue  # title, subtitle/subhead, logo, page-num, footer
            if _is_picture_idx_on_layout(slide, idx):
                continue  # a hero image on the other side of a Split Content
                          # layout starts at top=0 by design -- it's not in
                          # the subheading's column and can't "collide" with it
            if ph.top is None:
                continue
            if ph.top <= sub_ph.top:
                continue  # positioned above the subheading in this layout
                          # (e.g. a "Thank You" headline placeholder that
                          # sits above its tagline) -- not something a
                          # wrapped subheading could overlap downward into
            if not _x_ranges_overlap(sub_ph.left, sub_ph.width, ph.left, ph.width):
                continue  # different column entirely -- e.g. a contact-info
                          # block off to the side, not below in any real sense
            body_tops.append(ph.top)
        for shape in slide.shapes:
            name = getattr(shape, "name", "") or ""
            if not name.startswith("AFG_GEN_IDX"):
                continue
            shape_top = getattr(shape, "top", None)
            if shape_top is None or shape_top <= sub_ph.top:
                continue
            if not _x_ranges_overlap(sub_ph.left, sub_ph.width, getattr(shape, "left", None), getattr(shape, "width", None)):
                continue
            body_tops.append(shape_top)
        if not body_tops:
            continue
        body_top = min(body_tops)

        if est_sub_bottom > body_top:
            warnings.append(
                f"Slide {i}: subheading '{sub_text[:40]}' is likely to wrap to "
                f"{est_lines} lines at this placeholder's width, which will overlap "
                f"the content directly below it. Shorten the subheading -- "
                f"especially on narrower Split Content columns -- or verify "
                f"visually before sharing the file."
            )
    return warnings


def _detect_duplicate_body_content(prs):
    """
    Shared detection for verbatim-duplicate BODY content (bullets/text in
    content idx 14-21, never title/subtitle/footer/date/slide-number,
    which can legitimately repeat) appearing on 2+ different slides.
    Returns a list of (text, sorted slide-number list) tuples. Used both
    by check_duplicate_content (the full warning-list version, which also
    covers title/subtitle repeats as a softer signal) and by the hard
    block in build_deck/amend_deck that refuses to save a file containing
    one -- this happened in practice at real severity: the identical
    3-bullet block ("Reduce manual effort" / "Improve visibility" /
    "Enable mobile execution") appeared verbatim on 9 different
    topically-distinct slides after an amend that added new content,
    which is a content-generation bug (a reused generic filler pool
    instead of per-slide-specific writing), not a coincidence a human
    would ever produce.
    """
    seen = {}
    body_idxs = (14, 15, 16, 17, 18, 19, 20, 21)
    for i, slide in enumerate(prs.slides, start=1):
        for ph in slide.placeholders:
            if ph.placeholder_format.idx not in body_idxs or not ph.has_text_frame:
                continue
            for p in ph.text_frame.paragraphs:
                text = p.text.strip()
                if _word_count(text) >= 4:
                    seen.setdefault(text, []).append(i)
    return [(text, sorted(set(nums))) for text, nums in seen.items() if len(set(nums)) > 1]


def _block_on_duplicate_body_content(prs):
    """
    Hard stop before the file is saved: if any body-content line (a
    bullet, a diagram item's description, etc.) appears verbatim on 2+
    different slides, refuse to save and raise a PlanError naming the
    text and every affected slide. This exists because a warning can be
    skipped or ignored -- generic, reused filler content reaching the
    user is treated as a bug with zero acceptable frequency (the same
    standard blank slides are held to, though blank slides are now
    handled by auto-removal rather than a hard block -- see
    _remove_empty_slides -- since a slide can just be dropped cleanly,
    where duplicated content spread across otherwise-legitimate slides
    can't be "removed" the same way and needs a human to rewrite it).
    """
    dupes = _detect_duplicate_body_content(prs)
    if dupes:
        details = "; ".join(
            f"'{text[:50]}' on slides {nums}" for text, nums in dupes[:5]
        )
        more = f" (+{len(dupes) - 5} more)" if len(dupes) > 5 else ""
        raise PlanError(
            f"BLOCKED: found body content repeated verbatim across multiple "
            f"slides: {details}{more}. The file was NOT saved. This is the "
            f"'reused generic filler' failure mode, not a legitimate design "
            f"choice -- every slide's body content must be written "
            f"specifically for that slide's own topic. Rewrite the affected "
            f"slides with real, topic-specific content (never copy content "
            f"from one slide to another, even as a starting point) and "
            f"rebuild. This check cannot be bypassed; if two slides "
            f"genuinely need the exact same sentence for some specific "
            f"reason, that's rare enough to warrant asking the user rather "
            f"than assuming it here."
        )


def check_duplicate_content(prs):
    """
    Flags exact-duplicate strings (4+ words) appearing in placeholder text on
    more than one slide -- the "reused generic content pool" failure mode.
    This is a warning, not a hard error, since some repetition may be a
    deliberate, approved design choice (Section 4.1a).
    """
    seen = {}
    warnings = []
    for i, slide in enumerate(prs.slides, start=1):
        for ph in slide.placeholders:
            if not ph.has_text_frame:
                continue
            for p in ph.text_frame.paragraphs:
                text = p.text.strip()
                if _word_count(text) >= 4:
                    seen.setdefault(text, []).append(i)
    for text, slide_nums in seen.items():
        if len(set(slide_nums)) > 1:
            warnings.append(
                f"Duplicate content across slides {sorted(set(slide_nums))}: '{text[:60]}...' "
                f"-- confirm this repetition is intentional, otherwise rewrite per-slide."
            )
    return warnings


def check_chart_coverage(prs):
    """
    Flags a deck that contains zero native charts. Every deck should include
    at least one appropriate chart where the content supports it -- this is
    a warning (not a hard error) since a very short, purely qualitative
    update might genuinely have nothing chartable, but it should be a
    deliberate decision, not an oversight.
    """
    chart_count = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_chart:
                chart_count += 1
    if chart_count == 0:
        return [
            "This deck has zero charts. At least one appropriate chart is "
            "expected per deck -- if no slide has genuinely chartable content, "
            "that's fine, but confirm that's actually the case rather than "
            "defaulting to text-only."
        ]
    return []


def check_all_charts_illustrative(prs):
    """
    Flags a deck where every chart's title contains "(Illustrative)" --
    the marker used for fabricated placeholder numbers when no real data
    was available (Charts section: real user-provided figures first, then
    real published data via research, illustrative only as a last
    resort). A deck whose ONLY chart(s) are all illustrative usually means
    research was skipped rather than genuinely exhausted -- this happened
    in practice (a single illustrative line chart, no attempt at real
    data visible). Warning only, since a genuinely research-blocked topic
    can legitimately end up all-illustrative -- the point is to check
    that was a deliberate outcome, not the default path taken.
    """
    chart_shapes = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_chart:
                chart_shapes.append(shape.chart)
    if not chart_shapes:
        return []
    illustrative_count = 0
    for chart in chart_shapes:
        title = chart.chart_title.text_frame.text if chart.has_title else ""
        if "(Illustrative)" in title:
            illustrative_count += 1
    if illustrative_count == len(chart_shapes):
        return [
            f"Every chart in this deck ({illustrative_count}) is marked "
            f"'(Illustrative)' -- confirm real data was actually searched for "
            f"and genuinely unavailable, rather than defaulting straight to "
            f"illustrative figures. Real user-provided numbers or published "
            f"research (with a source cited) are always preferred; "
            f"illustrative is the last resort, not the default."
        ]
    return []


def check_no_reference_library_usage(prs):
    """
    Flags a deck of any real size (5+ slides, excluding Title Page and
    Thank You, which never need cloning) that contains ZERO shapes from
    either cloning mechanism (`clone_slide`'s AFG_GEN_WHOLESLIDE_ tag or
    `cloned_visual`'s AFG_GEN_IDX..._CLONEDVISUAL_ tag). This happened in
    practice, repeatedly, even after the 102-slide reference table was
    added directly to this skill's core instructions -- a deck built
    using only the base layouts, with the far larger reference
    library never touched at all. This is a coarse signal (it can't tell
    whether a genuine match existed for THIS deck's specific content),
    but zero usage across a normal-length deck is very rarely the
    correct outcome -- warn every time it happens so it's at least
    reconsidered rather than silently repeating the same miss.
    """
    if len(prs.slides) < 5:
        return []
    for slide in prs.slides:
        for shape in slide.shapes:
            name = getattr(shape, "name", "") or ""
            if name.startswith("AFG_GEN_WHOLESLIDE_") or "_CLONEDVISUAL_" in name:
                return []
    return [
        f"This deck has {len(prs.slides)} slides and uses ZERO cloned "
        f"reference-library slides (`clone_slide` or `cloned_visual`) -- "
        f"check the full 102-slide reference table again before treating "
        f"this as correct. Every one of the 102 reference slides is "
        f"equally clonable (no 'verified' subset), so a deck this size "
        f"using only the base layouts, with the reference library never "
        f"touched at all, is the exact failure mode this check exists to "
        f"catch -- it has happened before even with the table available."
    ]


def check_chart_variety(prs):
    """
    Flags a deck with 2+ charts that all use the identical chart_type
    (e.g. every chart is a bar chart) -- different data stories call for
    different chart types (a trend over time reads better as a line chart,
    a share-of-whole reads better as a pie chart, a comparison across
    categories reads well as a bar chart), and defaulting to bar every
    time regardless of the underlying data is a missed, easy improvement.
    Warning only -- if every chart in the deck genuinely is a category
    comparison, bar for all of them is correct; the point is to check the
    default wasn't just habit.
    """
    chart_types = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_chart:
                chart_types.append(shape.chart.chart_type)
    if len(chart_types) >= 2 and len(set(chart_types)) == 1:
        return [
            f"This deck has {len(chart_types)} charts, all the same type "
            f"({chart_types[0]}). Match chart type to what the data is "
            f"actually showing -- a trend over time reads better as a line "
            f"chart, a share of a whole as a pie chart, a comparison across "
            f"categories as a bar chart. Confirm this is genuinely the right "
            f"type for every one of them, not just the default."
        ]
    return []


def check_sparse_captions(prs):
    """
    Flags icon-paired captions on "Content Slide_4 columns" layouts that are
    too short to be a real caption (Section 2A: bare one/two-word labels
    like "Faster Delivery" sitting under an icon with nothing else on the
    slide are the main cause of half-empty-looking slides). Warning only --
    a short proper noun is occasionally the right call -- but it must be a
    deliberate choice, not the default.
    """
    warnings = []
    for i, slide in enumerate(prs.slides, start=1):
        family = _layout_family(slide.slide_layout.name)
        if family != "Content Slide_4 columns":
            continue
        ph_map = _placeholder_map(slide)
        for pic_idx, text_idx in FOUR_COLUMN_CAPTION_IDX.items():
            ph = ph_map.get(text_idx)
            if ph is None or not ph.has_text_frame:
                continue
            text = ph.text_frame.text.strip()
            if not text:
                continue
            if _word_count(text) < MIN_CAPTION_WORDS:
                warnings.append(
                    f"Slide {i} idx {text_idx}: caption '{text}' is only "
                    f"{_word_count(text)} word(s). Icon captions should be a short "
                    f"phrase explaining what it means or why it matters (e.g. "
                    f"'Cuts manual approval time in half'), not a bare label -- "
                    f"bare labels are the main cause of sparse-looking icon slides."
                )
    return warnings


def check_layout_rhythm(prs):
    """
    Flags a run of 4+ consecutive slides using the same layout "family"
    (e.g. every content slide being "Content Slide_N columns" back to back)
    -- a well-composed deck alternates text/icon slides against image-forward
    and chart slides (Section 2 rotation rule). Warning only, since a
    genuinely repetitive section (e.g. a long FAQ) can legitimately want
    this.
    """
    warnings = []
    families = [_layout_family(s.slide_layout.name) for s in prs.slides]
    run_start = 0
    for i in range(1, len(families) + 1):
        if i == len(families) or families[i] != families[run_start]:
            run_len = i - run_start
            if run_len > MAX_CONSECUTIVE_SAME_FAMILY:
                warnings.append(
                    f"Slides {run_start + 1}-{i}: {run_len} consecutive slides all "
                    f"use '{families[run_start]}' layout. Consider breaking up this "
                    f"run with an image-forward layout (Split Content image / "
                    f"Chapter Slide / Full image Slide) or a chart slide for rhythm."
                )
            run_start = i
    return warnings


# Maps a generated shape's tag-name "kind" component (see _tag_name) back to
# the diagram_type that produced it, so a saved deck can be inspected for
# which diagram types were actually used per slide without re-reading the
# plan JSON. Checked longest-prefix-first since "STEP_TEXT" and "STATVALUE"
# etc. share no ambiguous prefixes, but this keeps the lookup robust if more
# kinds are added later.
_DIAGRAM_KIND_TAGS = {
    "CARD": "card_grid",
    "BADGE": "process_steps",
    "STEP_TEXT": "process_steps",
    "ACCENTBAR": "accent_list",
    "ACCENT_TEXT": "accent_list",
    "STATVALUE": "stat_row",
    "STATBAR": "stat_row",
    "STATLABEL": "stat_row",
}


def _slide_diagram_kinds(slide):
    """Returns the set of diagram_type values present on this slide, inferred
    from generated-shape tag names (see _tag_name / _DIAGRAM_KIND_TAGS)."""
    prefix = "AFG_GEN_IDX"
    kinds = set()
    for shape in slide.shapes:
        name = getattr(shape, "name", "") or ""
        if not name.startswith(prefix):
            continue
        rest = name[len(prefix):]
        # rest looks like "<idx>_<KIND>_<n>" or "<idx>_<KIND>"
        after_idx = rest.split("_", 1)
        if len(after_idx) != 2:
            continue
        kind_part = after_idx[1]
        for tag, diagram_type in sorted(_DIAGRAM_KIND_TAGS.items(), key=lambda kv: -len(kv[0])):
            if kind_part.startswith(tag):
                kinds.add(diagram_type)
                break
    return kinds


# How many consecutive diagram-bearing slides may use the identical
# diagram_type before it reads as "the same illustration repeated" rather
# than a deliberate, varied composition.
MAX_CONSECUTIVE_SAME_DIAGRAM = 2


def _detect_empty_slides(prs):
    """
    Shared detection logic for "no real body content" slides -- returns a
    list of (1-based slide index, layout name) tuples. Used both by
    check_empty_slide (the warning-list version) and by the hard block in
    build_deck/amend_deck that refuses to save a file containing one (see
    those functions) -- a warning alone can be missed or ignored, and a
    blank slide reaching the user is a bug with no acceptable frequency,
    so this detection also gates the save itself, not just the warning list.
    """
    empty = []
    for i, slide in enumerate(prs.slides, start=1):
        family = _layout_family(slide.slide_layout.name)
        if family == "Title Page":
            continue  # title + subtitle only is the whole point of this layout, by design
        ph_map = _placeholder_map(slide)
        has_content = False
        for idx, ph in ph_map.items():
            if idx in (0, 1, 10, 11, 12, 13):
                continue  # title, subtitle/subhead, logo, page-num, footer
            if ph.has_text_frame and ph.text_frame.text.strip():
                has_content = True
                break
            if _is_picture_idx_on_layout(slide, idx) and not _is_picture_ph(ph):
                has_content = True  # picture placeholder already has an image in it
                break
        if not has_content:
            for shape in slide.shapes:
                name = getattr(shape, "name", "") or ""
                if name.startswith("AFG_GEN_IDX") or name.startswith("AFG_GEN_WHOLESLIDE_"):
                    has_content = True
                    break
                if getattr(shape, "has_chart", False):
                    has_content = True
                    break
        if not has_content:
            empty.append((i, slide.slide_layout.name))
    return empty


def _remove_empty_slides(prs, warnings):
    """
    Removes any slide with no real body content from the deck entirely,
    before it's ever saved -- rather than just blocking the save with an
    error. A blank slide reaching the user is treated as a bug with zero
    acceptable frequency; auto-removing it (with a loud warning explaining
    what was removed and why) guarantees that outcome unconditionally,
    without requiring a human to notice an error message and manually
    intervene before anything can be delivered. The warning still names
    every removed slide and its layout so the underlying planning gap
    (a dropped plan entry, a placeholder that never got filled) gets
    fixed for next time, not just silently swept away.

    Processes indices in reverse so removing one slide never shifts the
    position of another still pending removal, and drops the
    now-orphaned slide part's relationship (matching how delete_slides
    already cleans up) so the saved file has no dangling references.
    """
    empty = _detect_empty_slides(prs)
    if not empty:
        return
    xml_slides = prs.slides._sldIdLst
    slide_ids = list(xml_slides)
    for i, layout in sorted(empty, reverse=True):
        sldId = slide_ids[i - 1]
        xml_slides.remove(sldId)
        rId = sldId.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        prs.part.drop_rel(rId)
    details = "; ".join(f"slide {i} ({layout})" for i, layout in sorted(empty))
    warnings.append(
        f"REMOVED {len(empty)} blank slide(s) automatically before saving -- "
        f"no real body content was found on: {details}. A blank slide is "
        f"never acceptable output, so it was stripped rather than shipped, "
        f"but this still means a planning gap produced it (a dropped plan "
        f"entry, or a placeholder that never got filled) -- check whether "
        f"content was meant to be there and, if so, add it back properly "
        f"(don't just note that removal happened and move on)."
    )


def check_empty_slide(prs):
    """
    Flags a slide that has a title/subheading (or nothing at all) but no
    real body content -- no filled bullet/text placeholder, no image, no
    chart, no diagram. A content slide with nothing underneath its heading
    is always a mistake (an accidentally-skipped placeholder, a plan entry
    that got dropped, or similar) -- there's no legitimate reason to ship
    one, so this fires every time rather than only past a threshold.

    This warning is a secondary signal for humans reading the warning list
    -- the actual guarantee against a blank slide shipping is the hard
    block in build_deck/amend_deck (see _detect_empty_slides), which
    refuses to save the file at all rather than relying on this warning
    being read and acted on.
    """
    return [
        f"Slide {i}: no body content found -- only a title/subheading (or "
        f"nothing at all). This looks like an accidentally empty slide (a "
        f"dropped plan entry, a placeholder that never got filled). Add "
        f"content or remove the slide before sharing the file."
        for i, _layout_name in _detect_empty_slides(prs)
    ]


def check_chapter_before_thank_you(prs):
    """
    Flags a "Chapter Slide" (section divider -- image + big title +
    tagline, no bullets/data/diagram by design) sitting immediately
    before the closing "Thank you and contact" slide, with no real
    content slide in between. A chapter slide exists to introduce the
    section that follows it -- used as the second-to-last slide with
    nothing after it but Thank You, it introduces nothing and reads as a
    superfluous filler slide (a "vision statement" slide with no
    supporting content of its own). This happened in practice. This is
    NOT the same failure as check_empty_slide -- the chapter slide
    typically has real title/tagline text, so it isn't "empty" by that
    check's definition; the problem is structural placement, not missing
    content. Warning only, since a deliberate closing-chapter slide
    immediately before Thank You could rarely be intentional, but this
    combination is the pattern to double-check.
    """
    warnings = []
    families = [_layout_family(s.slide_layout.name) for s in prs.slides]
    for i in range(len(families) - 1):
        if families[i] == "Chapter Slide" and families[i + 1] == "Thank you and contact":
            warnings.append(
                f"Slide {i + 1}: a Chapter/divider slide sits directly before the "
                f"closing Thank You slide with no content slide in between -- it "
                f"introduces nothing and reads as filler. Either add real content "
                f"slides after it to justify the section break, fold its message "
                f"into the Thank You slide itself, or remove it."
            )
    return warnings


_ACCENT_ITEM_RE = re.compile(r"^AFG_GEN_IDX(\d+)_(ACCENT_ICON|ACCENTBAR)_\d+$")


def check_mixed_icon_usage(prs):
    """
    Flags an accent_list call where SOME items got a bundled icon and
    others got the plain accent bar instead -- e.g. item 1 and 3 have
    icons, item 2 doesn't. This reads as an accident (a forgotten "icon"
    field), not a deliberate style choice, and happened in practice.
    Within one accent_list call, every item should have an icon, or none
    should -- consistency is the rule, not "use icons where convenient."
    """
    warnings = []
    for i, slide in enumerate(prs.slides, start=1):
        idx_kinds = {}
        for shape in slide.shapes:
            m = _ACCENT_ITEM_RE.match(getattr(shape, "name", "") or "")
            if m:
                idx_kinds.setdefault(m.group(1), set()).add(m.group(2))
        for idx, kinds in idx_kinds.items():
            if len(kinds) > 1:
                warnings.append(
                    f"Slide {i} idx {idx}: this accent_list mixes icon items and "
                    f"plain-accent-bar items in the same call -- give ALL items an "
                    f"'icon', or omit 'icon' from all of them, so the list reads "
                    f"consistently. Mixing reads as a mistake, not a design choice."
                )
    return warnings


def check_mixed_bullet_styles_on_slide(prs):
    """
    Flags a slide where multiple "bullets" placeholders use different
    bullet_style values from each other -- e.g. one column uses "check"
    and the sibling column uses "arrow". Different slides can use
    different styles freely (that's the variety Section "Diagrams"/
    bullet-style guidance asks for), but WITHIN one slide the marker
    should be consistent -- two different glyphs side by side on the same
    slide reads as inconsistent formatting, not a deliberate choice. This
    happened in practice.
    """
    warnings = []
    for i, slide in enumerate(prs.slides, start=1):
        ph_map = _placeholder_map(slide)
        styles_seen = set()
        for idx, ph in ph_map.items():
            if idx in (0, 1, 10, 11, 12, 13) or not ph.has_text_frame:
                continue
            text = ph.text_frame.text
            if not text:
                continue
            first_char = text.lstrip()[:1]
            for style_name, glyph in BULLET_GLYPHS.items():
                if first_char == glyph:
                    styles_seen.add(style_name)
                    break
        if len(styles_seen) > 1:
            warnings.append(
                f"Slide {i}: multiple bullet styles used on the same slide "
                f"({sorted(styles_seen)}) -- pick one bullet_style per slide "
                f"and apply it consistently across every bulleted placeholder "
                f"on that slide. Different slides may use different styles."
            )
    return warnings


def check_generic_contact_info(prs):
    """
    Flags the "Thank you and contact" layout's contact block (idx16) when
    it doesn't look like real presenter contact info -- no email address
    pattern found at all. This happened in practice: idx16 held a generic
    description of the deck itself ("Executive Analytics Presentation")
    instead of the presenter's name/title/org/email. A closing slide's
    contact block exists specifically to say who to follow up with; a
    description of the presentation is never the right content for it,
    even as a placeholder pending real details.
    """
    warnings = []
    email_re = re.compile(r"[^\s@]+@[^\s@]+\.[^\s@]+")
    for i, slide in enumerate(prs.slides, start=1):
        if _layout_family(slide.slide_layout.name) != "Thank you and contact":
            continue
        ph_map = _placeholder_map(slide)
        contact_ph = ph_map.get(16)
        if contact_ph is None or not contact_ph.has_text_frame:
            continue
        text = contact_ph.text_frame.text.strip()
        if not text:
            continue  # separately covered by check_empty_slide if the whole slide is bare
        if not email_re.search(text):
            warnings.append(
                f"Slide {i}: the contact block ('{text[:50]}') doesn't contain an "
                f"email address -- it should hold the actual presenter's name, "
                f"title, organization, and email (e.g. 'Jane Doe / Power Platform "
                f"Lead / AFG Digital / jane.doe@afg.com'), not a description of "
                f"the presentation itself. Ask for the presenter's real details "
                f"if they weren't given rather than inventing a generic line."
            )
    return warnings


def check_narrow_diagram_full_width_layout(prs):
    """
    Flags a process_steps or accent_list diagram placed on "Content Slide_1
    column" -- a full-width layout with no image slot at all. Both of
    those diagram types are naturally narrow content (a numbered list or
    an accent list doesn't benefit from being stretched wide), so on a
    layout with nothing else to fill the remaining width, a large portion
    of the slide reads as empty background. This happened in practice.
    There's no in-place fix (the layout itself has no second slot to add
    an image to), so this always warns rather than only warning past some
    threshold -- the fix is a different layout choice, made before the
    next build/amend, not something this script can patch afterward.
    """
    warnings = []
    NARROW_TYPES = {"process_steps", "accent_list"}
    for i, slide in enumerate(prs.slides, start=1):
        family = _layout_family(slide.slide_layout.name)
        if family != "Content Slide_1 column":
            continue
        kinds = _slide_diagram_kinds(slide)
        hit = kinds & NARROW_TYPES
        if hit:
            warnings.append(
                f"Slide {i}: a '{sorted(hit)[0]}' diagram is on 'Content Slide_1 "
                f"column', a full-width layout with no image slot -- this content is "
                f"naturally narrow, so the unused width will read as empty background. "
                f"Use a Split Content layout (text + image) instead and pair this "
                f"diagram with a relevant photo/illustration, or switch to a "
                f"full-width-friendly type (multi-item card_grid, or stat_row) if "
                f"there's genuinely no image to use."
            )
    return warnings


def check_diagram_variety(prs):
    """
    Flags a run of more than MAX_CONSECUTIVE_SAME_DIAGRAM diagram-bearing
    slides that all use the SAME diagram_type (e.g. four "card_grid" slides
    in a row, even if separated by non-diagram slides in between isn't
    relevant here -- this checks slides that actually contain a diagram,
    filtering out slides with no diagram at all, since those don't
    contribute to "everything looks the same"). Warning only.
    """
    warnings = []
    diagram_slides = []
    for i, slide in enumerate(prs.slides, start=1):
        kinds = _slide_diagram_kinds(slide)
        if len(kinds) == 1:
            diagram_slides.append((i, next(iter(kinds))))
        elif len(kinds) > 1:
            diagram_slides.append((i, "mixed"))
    run_start = 0
    for i in range(1, len(diagram_slides) + 1):
        if i == len(diagram_slides) or diagram_slides[i][1] != diagram_slides[run_start][1]:
            run_len = i - run_start
            kind = diagram_slides[run_start][1]
            if run_len > MAX_CONSECUTIVE_SAME_DIAGRAM and kind != "mixed":
                slide_nums = [s[0] for s in diagram_slides[run_start:i]]
                warnings.append(
                    f"Slides {slide_nums}: {run_len} diagram slides in a row all use "
                    f"'{kind}' -- vary the diagram type (process_steps / card_grid / "
                    f"accent_list / stat_row) so the deck doesn't read as the same "
                    f"illustration repeated. See SKILL.md 'Diagrams' variety guidance."
                )
            run_start = i
    return warnings


def check_sparse_columns(prs):
    """
    Flags a "Content Slide_N columns" (or the single-column full-width
    "Content Slide_1 column") slide where every content placeholder is
    still plain bulleted text (never converted to a diagram) and each
    holds only a handful of short bullets -- e.g. three bare columns of 3
    short items each, or a single full-width column with 4 short bullets,
    with nothing else on the slide. This is the exact "half the slide is
    empty" pattern: short bullet lists sitting in the top portion of the
    slide with no chart/image/diagram using the rest of the space.
    Warning only, since a genuinely dense column (long bullets, many of
    them) is a legitimate use of the plain layout.
    """
    warnings = []
    CONTENT_IDXS = (14, 15, 16, 17, 19, 21)
    for i, slide in enumerate(prs.slides, start=1):
        family = _layout_family(slide.slide_layout.name)
        is_multi_column = "columns" in family and family != "Content Slide_4 columns"
        is_single_column = family == "Content Slide_1 column"
        if not (is_multi_column or is_single_column):
            continue  # 4-column icon layout has its own caption check
        ph_map = _placeholder_map(slide)
        content_phs = [ph_map[idx] for idx in CONTENT_IDXS if idx in ph_map]
        if not content_phs:
            continue  # already converted to a diagram (placeholder was deleted)
        all_sparse = True
        for ph in content_phs:
            if not ph.has_text_frame:
                all_sparse = False
                break
            n_paras = len([p for p in ph.text_frame.paragraphs if p.text.strip()])
            if n_paras == 0 or n_paras > SPARSE_COLUMN_MAX_BULLETS:
                all_sparse = False
                break
        if all_sparse and is_single_column:
            warnings.append(
                f"Slide {i}: 'Content Slide_1 column' (full-width, no image slot) "
                f"filled with only a short plain bullet list and nothing else -- this "
                f"leaves most of the slide empty both vertically and horizontally. "
                f"Pair this content with an image via a Split Content layout instead, "
                f"or convert to a full-height diagram ('process_steps' for a sequence, "
                f"'accent_list' otherwise, both need 2+ items) if there's genuinely no "
                f"image to use."
            )
        elif all_sparse:
            warnings.append(
                f"Slide {i}: '{family}' filled with plain short bullet lists in "
                f"every column and nothing else on the slide -- this tends to leave "
                f"much of the slide height empty. Consider 'type': 'diagram', "
                f"'diagram_type': 'card_grid' with a 'bullets' list per item instead "
                f"(one card per category, e.g. 'Retail', 'HR', 'Operations', each "
                f"listing its own short items) -- see SKILL.md 'Diagrams' section."
            )
    return warnings


def summarize(prs, output_path, warnings):
    lines = [f"Saved: {output_path} ({len(prs.slides)} slides)"]
    for i, slide in enumerate(prs.slides, start=1):
        layout_name = slide.slide_layout.name
        title = ""
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 0 and ph.has_text_frame:
                title = ph.text_frame.text.strip()
                break
        lines.append(f"  Slide {i}: {layout_name} -- \"{title}\"")
    if warnings:
        lines.append("Warnings:")
        for w in warnings:
            lines.append(f"  - {w}")
    return "\n".join(lines)


def inspect_deck(input_path):
    prs = Presentation(input_path)
    lines = [f"{input_path}: {len(prs.slides)} slides"]
    for i, slide in enumerate(prs.slides, start=1):
        lines.append(f"Slide {i}: layout='{slide.slide_layout.name}'")
        placeholder_ids = set()
        for ph in slide.placeholders:
            placeholder_ids.add(ph.shape_id)
            kind = "PIC" if _is_picture_ph(ph) else "TEXT"
            text = ph.text_frame.text.strip().replace("\n", " | ") if ph.has_text_frame else "<image>"
            lines.append(f"    idx={ph.placeholder_format.idx} [{kind}] {text[:80]}")
        # Report non-placeholder shapes too (e.g. charts, which replace their
        # original placeholder and so won't show up in slide.placeholders).
        for shape in slide.shapes:
            if shape.shape_id in placeholder_ids:
                continue
            if shape.has_chart:
                lines.append(f"    [CHART] '{shape.chart.chart_title.text_frame.text if shape.chart.has_title else '(untitled)'}' at ({shape.left},{shape.top})")
            elif shape.shape_type is not None and "PICTURE" in str(shape.shape_type):
                lines.append(f"    [IMAGE] at ({shape.left},{shape.top})")
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="AFG deck builder (python-pptx, idx-exact)")
    sub = parser.add_subparsers(dest="command", required=True)

    p_build = sub.add_parser("build")
    p_build.add_argument("--template", required=True)
    p_build.add_argument("--plan", required=True)
    p_build.add_argument("--output", required=True)

    p_amend = sub.add_parser("amend")
    p_amend.add_argument("--input", required=True)
    p_amend.add_argument("--plan", required=True)
    p_amend.add_argument("--output", required=True)
    p_amend.add_argument("--reference-template", required=False, default=None,
                          help="Path to the original AFG_template.pptx -- required only if the amend plan uses 'cloned_visual'.")

    p_inspect = sub.add_parser("inspect")
    p_inspect.add_argument("--input", required=True)

    p_catalog = sub.add_parser("catalog")
    p_catalog.add_argument("--template", required=True)
    p_catalog.add_argument("--slide", required=True, type=int,
                            help="1-based slide number in the original template to catalog.")

    args = parser.parse_args()

    try:
        if args.command == "build":
            with open(args.plan) as f:
                plan = json.load(f)
            print(build_deck(args.template, plan, args.output))
        elif args.command == "amend":
            with open(args.plan) as f:
                plan = json.load(f)
            print(amend_deck(args.input, plan, args.output, reference_template_path=args.reference_template))
        elif args.command == "inspect":
            print(inspect_deck(args.input))
        elif args.command == "catalog":
            print(catalog_reference_slide(args.template, args.slide))
    except PlanError as e:
        print(f"PLAN ERROR: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
