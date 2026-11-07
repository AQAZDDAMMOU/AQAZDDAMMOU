"""Structural diagram builders: architecture-layers, swimlane-process,
org-chart, value-chain.

Registered into builders.LAYOUT_MAP at import time (see bottom of
builders.py).  All geometry is design inches on the 13.33 x 7.5 canvas;
builders' primitives rescale to the actual slide size.  Content must clear
the footer zone (nothing below y = 7.0)."""
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

import builders as B
import shape_groups
from helpers import warn

MAX_LAYERS = 6
MAX_LAYER_ITEMS = 5
MAX_LANES = 4
MAX_LANE_STEPS = 5
MAX_ORG_NODES = 12
MAX_CHAIN_STEPS = 6


# ── architecture layers (target operating model / stack diagram) ─────────────
def build_architecture_layers_slide(prs, p, pal, ctx):
    """Horizontal stacked bands, top layer first.  Row syntax:
    '- Layer: Name="Experience" Items="Portal, Mobile" [Accent="2"]'."""
    layers = [layer for layer in p.get("layers", []) if layer.get("name", "").strip()]
    layers = layers[:MAX_LAYERS]
    if not layers:
        warn("architecture-layers has no '- Layer:' rows; rendering bullet-list")
        return B.build_bullet_slide(prs, p, pal, ctx)
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)

    top, bottom, gap, label_w = 1.9, 6.6, 0.12, 2.3
    n = len(layers)
    band_h = (bottom - top - gap * (n - 1)) / n
    for i, layer in enumerate(layers):
        y = top + i * (band_h + gap)
        accent = pal["accent2"] if layer.get("accent") == "2" else pal["accent1"]
        B.add_rect(slide, 0.7, y, label_w, band_h, accent)
        B.add_tb(slide, layer["name"], 0.86, y + band_h / 2 - 0.24,
                 label_w - 0.32, 0.48, size=14, bold=True,
                 color=pal["bg_deep"], font=pal["font_body"])
        B.add_rect(slide, 0.7 + label_w + 0.08, y, 11.9 - label_w - 0.08,
                   band_h, pal["surface"])
        items = [s.strip() for s in layer.get("items", "").split(",")
                 if s.strip()][:MAX_LAYER_ITEMS]
        if not items:
            continue
        chip_zone_x = 0.7 + label_w + 0.24
        chip_zone_w = 11.9 - label_w - 0.40
        cw = chip_zone_w / len(items)
        for j, item in enumerate(items):
            cx = chip_zone_x + j * cw
            B.add_rect(slide, cx, y + 0.14, cw - 0.16, band_h - 0.28,
                       pal["bg"], line_hex=pal["text_muted"], line_pt=0.75)
            B.add_tb(slide, item, cx + 0.08, y + band_h / 2 - 0.22,
                     cw - 0.32, 0.44, size=11, color=pal["text"],
                     align=PP_ALIGN.CENTER, font=pal["font_label"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


# ── swimlane process (lane-per-actor flow) ───────────────────────────────────
def _lane_steps(lane):
    return [s.strip() for s in lane.get("steps", "").split(">")
            if s.strip()][:MAX_LANE_STEPS]


def build_swimlane_slide(prs, p, pal, ctx):
    """Lane-per-actor process.  Row syntax:
    '- Lane: Name="Business" Steps="Submit idea > Refine case"'.
    Steps flow left-to-right inside each lane with arrow connectors."""
    lanes = [lane for lane in p.get("lanes", []) if lane.get("name", "").strip()]
    lanes = lanes[:MAX_LANES]
    if not lanes or not any(_lane_steps(lane) for lane in lanes):
        warn("swimlane-process has no '- Lane:' rows with Steps; "
             "rendering bullet-list")
        return B.build_bullet_slide(prs, p, pal, ctx)
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)

    top, bottom, label_w = 1.9, 6.7, 1.9
    n_cols = max(len(_lane_steps(lane)) for lane in lanes)
    lane_h = (bottom - top) / len(lanes)
    col_w = (11.9 - label_w) / n_cols
    for i, lane in enumerate(lanes):
        y = top + i * lane_h
        band = pal["surface"] if i % 2 == 0 else pal["bg"]
        B.add_rect(slide, 0.7, y, 11.9, lane_h - 0.06, band)
        B.add_rect(slide, 0.7, y, 0.06, lane_h - 0.06, pal["accent1"])
        B.add_tb(slide, lane["name"], 0.88, y + lane_h / 2 - 0.25,
                 label_w - 0.3, 0.5, size=13, bold=True, color=pal["text"],
                 font=pal["font_body"])
        steps = _lane_steps(lane)
        box_h = min(0.9, lane_h - 0.34)
        for j, step in enumerate(steps):
            x = 0.7 + label_w + j * col_w + 0.10
            by = y + (lane_h - 0.06 - box_h) / 2
            B.add_rect(slide, x, by, col_w - 0.44, box_h, pal["bg_deep"],
                       line_hex=pal["accent1"], line_pt=1.0)
            B.add_tb(slide, step, x + 0.08, by + box_h / 2 - 0.28,
                     col_w - 0.60, 0.56, size=11, color=pal["text"],
                     align=PP_ALIGN.CENTER, font=pal["font_label"])
            if j:  # arrow from previous box in this lane
                ax = x - 0.34
                ax_l, ax_t, ax_w, ax_h = B._sc(
                    ax, y + lane_h / 2 - 0.14, 0.24, 0.20)
                arrow = slide.shapes.add_shape(
                    33, Inches(ax_l), Inches(ax_t), Inches(ax_w), Inches(ax_h)
                )  # 33 = RIGHT_ARROW
                arrow.fill.solid()
                from palettes import hex_rgb
                arrow.fill.fore_color.rgb = hex_rgb(pal["accent2"])
                arrow.line.fill.background()
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


# ── org chart / governance model (top-down tree) ─────────────────────────────
def build_org_chart_slide(prs, p, pal, ctx):
    """Top-down hierarchy.  Row syntax:
    '- Node: Id="coe" Label="AI CoE" Parent="steer" [Role="Delivery"]'.
    Root(s) at top; boxes show Label bold + optional Role muted."""
    nodes = [n for n in p.get("nodes", []) if n.get("id", "").strip()]
    nodes = nodes[:MAX_ORG_NODES]
    if not nodes:
        warn("org-chart needs '- Node:' rows; rendering bullet-list")
        return B.build_bullet_slide(prs, p, pal, ctx)
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)

    by_id = {n["id"]: n for n in nodes}
    children = {}
    for n in nodes:
        par = n.get("parent", "").strip()
        if par and par in by_id:
            children.setdefault(par, []).append(n["id"])
    roots = [n["id"] for n in nodes if not n.get("parent", "").strip()]
    root = roots[0] if roots else nodes[0]["id"]

    depths, queue = {root: 1}, [root]
    while queue:
        nid = queue.pop(0)
        for ch in children.get(nid, []):
            if ch not in depths:
                depths[ch] = depths[nid] + 1
                queue.append(ch)
    skipped = sorted(i for i in by_id if i not in depths)
    if skipped:
        warn("org-chart: nodes unreachable from root skipped: "
             + ", ".join(skipped))
    max_d = max(depths.values())

    top, bottom = 1.95, 6.6
    row_h = (bottom - top) / max_d
    box_h = max(0.55, min(1.0, row_h - 0.45))

    def _leaves(nid, seen):
        if nid in seen:
            return []
        seen.add(nid)
        kids = [k for k in children.get(nid, []) if k in depths]
        if not kids:
            return [nid]
        out = []
        for k in kids:
            out.extend(_leaves(k, seen))
        return out

    leaves = _leaves(root, set())
    col_w = 11.9 / max(len(leaves), 1)
    box_w = min(2.8, col_w - 0.15)
    xs = {nid: 0.7 + i * col_w + col_w / 2 for i, nid in enumerate(leaves)}

    def center_x(nid):
        if nid not in xs:
            kids = [k for k in children[nid] if k in depths]
            xs[nid] = sum(center_x(k) for k in kids) / len(kids)
        return xs[nid]
    center_x(root)

    for nid, d in depths.items():
        node = by_id[nid]
        x = xs[nid] - box_w / 2
        y = top + (d - 1) * row_h
        fill = pal["accent1"] if d == 1 else pal["surface"]
        B.add_rect(slide, x, y, box_w, box_h, fill)
        B.add_tb(slide, node.get("label", nid), x + 0.06, y + 0.05,
                 box_w - 0.12, 0.42, size=12, bold=True,
                 color=pal["bg_deep"] if d == 1 else pal["text"],
                 align=PP_ALIGN.CENTER, font=pal["font_body"])
        if node.get("role"):
            B.add_tb(slide, node["role"], x + 0.06, y + box_h - 0.36,
                     box_w - 0.12, 0.32, size=10,
                     color=pal["bg_deep"] if d == 1 else pal["text_muted"],
                     align=PP_ALIGN.CENTER, font=pal["font_label"])

    from palettes import hex_rgb
    for par, kids in children.items():
        if par not in depths:
            continue
        py = top + (depths[par] - 1) * row_h + box_h
        for kid in kids:
            if kid not in depths:
                continue
            ky = top + (depths[kid] - 1) * row_h
            x1, y1, _, _ = B._sc(xs[par], py, 0, 0)
            x2, y2, _, _ = B._sc(xs[kid], ky, 0, 0)
            conn = slide.shapes.add_connector(
                MSO_CONNECTOR.ELBOW, Inches(x1), Inches(y1),
                Inches(x2), Inches(y2))
            conn.line.color.rgb = hex_rgb(pal["text_muted"])
            conn.line.width = Pt(1.25)
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


# ── value chain (chevron sequence + support band) ────────────────────────────
def build_value_chain_slide(prs, p, pal, ctx):
    """Porter-style chain: chevron per '- Step N: Title=".." Sub="a, b"',
    sub-activities listed under each chevron, optional '- Support: "..."'
    band across the bottom."""
    steps = p.get("steps", [])[:MAX_CHAIN_STEPS]
    if len(steps) < 2:
        warn("value-chain needs 2+ '- Step N:' rows; rendering bullet-list")
        return B.build_bullet_slide(prs, p, pal, ctx)
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)

    from palettes import hex_rgb
    n = len(steps)
    gap = 0.10
    w = (11.9 - gap * (n - 1)) / n
    chev_y, chev_h = 2.15, 1.05
    for i, step in enumerate(steps):
        x = 0.7 + i * (w + gap)
        chev_l, chev_t, chev_w, chev_hh = B._sc(x, chev_y, w, chev_h)
        chev = slide.shapes.add_shape(52, Inches(chev_l), Inches(chev_t),
                                      Inches(chev_w), Inches(chev_hh))  # 52 = CHEVRON
        chev.fill.solid()
        chev.fill.fore_color.rgb = hex_rgb(
            pal["accent1"] if i == 0 else pal["surface"])
        chev.line.fill.background()
        B.add_tb(slide, step.get("title", ""), x + 0.30, chev_y + 0.28,
                 w - 0.42, 0.5, size=13, bold=True,
                 color=pal["bg_deep"] if i == 0 else pal["text"],
                 align=PP_ALIGN.CENTER, font=pal["font_body"])
        if step.get("sub"):
            B.add_tb(slide, step["sub"], x + 0.05, chev_y + chev_h + 0.25,
                     w - 0.10, 2.3, size=11, color=pal["text_muted"],
                     font=pal["font_label"])

    if p.get("support"):
        sup_y = 6.0
        B.add_rect(slide, 0.7, sup_y, 11.9, 0.6, pal["bg_deep"],
                   line_hex=pal["accent2"], line_pt=1.0)
        B.add_tb(slide, "Enabling: " + p["support"], 0.9, sup_y + 0.12,
                 11.5, 0.4, size=12, color=pal["text"],
                 align=PP_ALIGN.CENTER, font=pal["font_body"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


LAYOUTS = {
    "architecture-layers": build_architecture_layers_slide,
    "swimlane-process": build_swimlane_slide,
    "org-chart": build_org_chart_slide,
    "value-chain": build_value_chain_slide,
}
