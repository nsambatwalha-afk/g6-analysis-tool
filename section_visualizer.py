"""
section_visualizer.py
Generates cross-section diagrams for steel sections used in the design tasks.
"""

import matplotlib
matplotlib.use("Agg")  # Non-interactive backend suitable for Streamlit

import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np
import openpyxl


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _lookup_i_section_dims(xlsx_path, designation):
    """
    Return (h, b, tw, tf) in mm for an I/H section from the given Excel file.
    Columns used (1-based): 1=designation, 15=tw, 16=h, 17=tf, 19=b.
    Returns None if designation is not found.
    """
    try:
        ws = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True).active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if str(row[0]) == str(designation):
                h  = float(row[15])   # col 16 → index 15
                tf = float(row[16])   # col 17 → index 16
                b  = float(row[18])   # col 19 → index 18
                tw = float(row[14])   # col 15 → index 14
                return h, b, tw, tf
    except Exception:
        pass
    return None


def _parse_chs_size(size_str):
    """
    Parse a string like 'CHS 219.1x5' or '219.1x5' into (dia, thickness).
    Returns (None, None) on failure.
    """
    try:
        cleaned = size_str.replace("CHS", "").strip()
        parts = cleaned.split("x")
        return float(parts[0]), float(parts[1])
    except Exception:
        return None, None


def _parse_angle_size(size_str):
    """
    Parse angle designation like '200x200' or '150x100' into (leg1, leg2).
    Returns (100, 100) as fallback.
    """
    try:
        cleaned = str(size_str).replace("L", "").replace(" ", "")
        parts = cleaned.split("x")
        leg1 = float(parts[0])
        leg2 = float(parts[1]) if len(parts) > 1 else leg1
        return leg1, leg2
    except Exception:
        return 100.0, 100.0


# ---------------------------------------------------------------------------
# Core drawing functions
# ---------------------------------------------------------------------------

def draw_i_section(h, b, tw, tf, title="I-Section", grade="", info_text=""):
    """
    Draw an I/H section cross-section.

    Parameters
    ----------
    h, b, tw, tf : floats (mm) – depth, flange width, web thickness, flange thickness
    title        : plot title (designation)
    grade        : steel grade string (shown in subtitle)
    info_text    : additional info line (e.g. member type, utilization)

    Returns
    -------
    matplotlib.figure.Figure
    """
    fig, ax = plt.subplots(figsize=(4.5, 5.5))
    ax.set_aspect("equal")

    hw = h - 2 * tf           # clear web height
    half_b  = b  / 2.0
    half_tw = tw / 2.0

    # ----- patches -----
    bottom_fl = patches.Rectangle((-half_b, -h / 2), b,  tf, lw=1, ec="black", fc="#4A90D9", alpha=0.85)
    top_fl    = patches.Rectangle((-half_b,  h / 2 - tf), b, tf, lw=1, ec="black", fc="#4A90D9", alpha=0.85)
    web       = patches.Rectangle((-half_tw, -h / 2 + tf), tw, hw, lw=1, ec="black", fc="#4A90D9", alpha=0.85)

    for p in (bottom_fl, top_fl, web):
        ax.add_patch(p)

    # ----- dimension annotations -----
    margin = max(b, h) * 0.18

    # height arrow (right side)
    ax.annotate("", xy=(half_b + margin, h / 2), xytext=(half_b + margin, -h / 2),
                arrowprops=dict(arrowstyle="<->", color="dimgray", lw=1.0))
    ax.text(half_b + margin + max(b, h) * 0.05, 0, f"h = {h:.0f} mm",
            va="center", ha="left", fontsize=7.5, color="dimgray")

    # width arrow (bottom)
    ax.annotate("", xy=(half_b, -h / 2 - margin), xytext=(-half_b, -h / 2 - margin),
                arrowprops=dict(arrowstyle="<->", color="dimgray", lw=1.0))
    ax.text(0, -h / 2 - margin * 1.9, f"b = {b:.0f} mm",
            ha="center", va="top", fontsize=7.5, color="dimgray")

    # tf label (inside top flange)
    ax.text(0, h / 2 - tf / 2, f"tf = {tf:.1f}", ha="center", va="center",
            fontsize=6.5, color="white", fontweight="bold")

    # tw label (inside web)
    ax.text(0, 0, f"tw\n{tw:.1f}", ha="center", va="center",
            fontsize=6.0, color="white", fontweight="bold")

    padding = max(b, h) * 0.55
    ax.set_xlim(-half_b - padding, half_b + padding)
    ax.set_ylim(-h / 2 - padding, h / 2 + padding)
    ax.axis("off")

    subtitle = f"Grade: {grade}" if grade else ""
    if info_text:
        subtitle = (subtitle + "\n" + info_text).strip("\n")
    ax.set_title(f"{title}\n{subtitle}", fontsize=10, fontweight="bold", pad=8)

    fig.tight_layout()
    return fig


def draw_chs(outer_dia, thickness, title="CHS", grade="", info_text=""):
    """
    Draw a Circular Hollow Section.

    Parameters
    ----------
    outer_dia : float (mm) – outer diameter
    thickness : float (mm) – wall thickness
    title     : plot title
    grade     : steel grade string
    info_text : additional info line

    Returns
    -------
    matplotlib.figure.Figure
    """
    fig, ax = plt.subplots(figsize=(4.5, 5.0))
    ax.set_aspect("equal")

    r_out = outer_dia / 2.0
    r_in  = r_out - thickness

    outer = plt.Circle((0, 0), r_out, lw=1.5, ec="black", fc="#4A90D9", alpha=0.85)
    inner = plt.Circle((0, 0), r_in,  lw=1.0, ec="black", fc="white")
    ax.add_patch(outer)
    ax.add_patch(inner)

    # Diameter arrow (top)
    margin = r_out * 0.35
    ax.annotate("", xy=(r_out, r_out + margin), xytext=(-r_out, r_out + margin),
                arrowprops=dict(arrowstyle="<->", color="dimgray", lw=1.0))
    ax.text(0, r_out + margin * 1.8, f"D = {outer_dia:.1f} mm",
            ha="center", va="bottom", fontsize=8, color="dimgray")

    # Thickness arrow (horizontal, through the wall)
    mid_r = (r_out + r_in) / 2.0
    ax.annotate("", xy=(r_out, 0), xytext=(r_in, 0),
                arrowprops=dict(arrowstyle="<->", color="darkorange", lw=1.2))
    ax.text(mid_r, -r_out * 0.12, f"t = {thickness:.1f} mm",
            ha="center", va="top", fontsize=7.5, color="darkorange")

    padding = r_out * 0.65
    ax.set_xlim(-r_out - padding, r_out + padding)
    ax.set_ylim(-r_out - padding, r_out + padding)
    ax.axis("off")

    subtitle = f"Grade: {grade}" if grade else ""
    if info_text:
        subtitle = (subtitle + "\n" + info_text).strip("\n")
    ax.set_title(f"{title}\n{subtitle}", fontsize=10, fontweight="bold", pad=8)

    fig.tight_layout()
    return fig


def draw_angle(leg1, leg2, thickness, title="Angle Section", grade="", info_text=""):
    """
    Draw an angle (L) section.

    Parameters
    ----------
    leg1, leg2 : floats (mm) – horizontal and vertical leg lengths
    thickness  : float (mm) – leg thickness
    title      : plot title
    grade      : steel grade string
    info_text  : additional info line

    Returns
    -------
    matplotlib.figure.Figure
    """
    fig, ax = plt.subplots(figsize=(4.5, 5.0))
    ax.set_aspect("equal")

    # Vertical leg (left)
    vert = patches.Rectangle((0, 0), thickness, leg2,
                              lw=1.5, ec="black", fc="#4A90D9", alpha=0.85)
    # Horizontal leg (bottom)
    horiz = patches.Rectangle((0, 0), leg1, thickness,
                               lw=1.5, ec="black", fc="#4A90D9", alpha=0.85)
    ax.add_patch(vert)
    ax.add_patch(horiz)

    margin = max(leg1, leg2) * 0.22

    # horizontal leg arrow
    ax.annotate("", xy=(leg1, -margin), xytext=(0, -margin),
                arrowprops=dict(arrowstyle="<->", color="dimgray", lw=1.0))
    ax.text(leg1 / 2, -margin * 1.9, f"b = {leg1:.0f} mm",
            ha="center", va="top", fontsize=8, color="dimgray")

    # vertical leg arrow
    ax.annotate("", xy=(-margin, leg2), xytext=(-margin, 0),
                arrowprops=dict(arrowstyle="<->", color="dimgray", lw=1.0))
    ax.text(-margin * 2.8, leg2 / 2, f"h = {leg2:.0f} mm",
            ha="center", va="center", fontsize=8, color="dimgray", rotation=90)

    # thickness label inside the corner block
    ax.text(thickness / 2, thickness / 2, f"t={thickness:.1f}",
            ha="center", va="center", fontsize=7, color="white", fontweight="bold")

    padding = max(leg1, leg2) * 0.55
    ax.set_xlim(-padding, leg1 + padding * 0.4)
    ax.set_ylim(-padding, leg2 + padding * 0.4)
    ax.axis("off")

    subtitle = f"Grade: {grade}" if grade else ""
    if info_text:
        subtitle = (subtitle + "\n" + info_text).strip("\n")
    ax.set_title(f"{title}\n{subtitle}", fontsize=10, fontweight="bold", pad=8)

    fig.tight_layout()
    return fig


# ---------------------------------------------------------------------------
# High-level helpers called from app.py
# ---------------------------------------------------------------------------

def visualize_tension_section(shapeten, section, grade="", info_text=""):
    """
    Return a Figure for a tension member section.

    section : tuple returned by truss_analysis.ten_designer()
              CHS   → (area, dia, thickness, rg)
              Angle → (area, size_str, thickness)
    """
    if shapeten == "CHS":
        dia = float(section[1])
        t   = float(section[2])
        return draw_chs(dia, t,
                        title=f"Tension Member — CHS {dia:.1f}×{t:.1f}",
                        grade=grade, info_text=info_text)

    elif shapeten == "Angle":
        size_str = section[1]
        t        = float(section[2])
        leg1, leg2 = _parse_angle_size(size_str)
        return draw_angle(leg1, leg2, t,
                          title=f"Tension Member — Angle {size_str}",
                          grade=grade, info_text=info_text)
    return None


def visualize_compression_section(shapecomp, section, grade="", info_text=""):
    """
    Return a Figure for a compression member section.

    section : dict returned by truss_analysis.comp_designer()
    """
    if shapecomp == "CHS":
        dia, t = _parse_chs_size(section["size"])
        if dia is None:
            return None
        return draw_chs(dia, t,
                        title=f"Compression Member — {section['size']}",
                        grade=grade, info_text=info_text)

    elif shapecomp in ("UB", "UC"):
        designation = section["size"]
        xlsx = "UB-2.xlsx" if shapecomp == "UB" else "UC-2.xlsx"
        dims = _lookup_i_section_dims(xlsx, designation)
        if dims is None:
            return None
        h, b, tw, tf = dims
        return draw_i_section(h, b, tw, tf,
                              title=f"Compression Member — {designation}",
                              grade=grade, info_text=info_text)

    elif shapecomp == "Angle":
        # comp_designer may return an Angle section dict with "size"
        size_str = section.get("size", "100x100")
        t = section.get("thickness", 10.0)
        leg1, leg2 = _parse_angle_size(size_str)
        return draw_angle(leg1, leg2, t,
                          title=f"Compression Member — Angle {size_str}",
                          grade=grade, info_text=info_text)
    return None


def visualize_beam_section(designation, grade="", beam_type="Beam", info_text=""):
    """
    Return a Figure for a UB beam section.

    designation : UB section designation string
    beam_type   : label shown in the title
    """
    dims = _lookup_i_section_dims("UB-2.xlsx", designation)
    if dims is None:
        return None
    h, b, tw, tf = dims
    return draw_i_section(h, b, tw, tf,
                          title=f"{beam_type} — {designation}",
                          grade=grade, info_text=info_text)


def visualize_beam_column_section(designation, shape, grade="", info_text=""):
    """
    Return a Figure for a UB or UC beam-column section.

    designation : section designation string
    shape       : "UB" or "UC"
    """
    xlsx = "UB-2.xlsx" if shape == "UB" else "UC-2.xlsx"
    dims = _lookup_i_section_dims(xlsx, designation)
    if dims is None:
        return None
    h, b, tw, tf = dims
    return draw_i_section(h, b, tw, tf,
                          title=f"Beam-Column — {designation}",
                          grade=grade, info_text=info_text)
