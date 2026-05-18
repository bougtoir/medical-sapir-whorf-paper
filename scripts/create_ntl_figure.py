#!/usr/bin/env python3
"""
Generate the Nosological Therapeutic Loop (NTL) diagram as PNG images.
Output: medical_sapir_whorf/output/figure1_ntl_en.png
        medical_sapir_whorf/output/figure1_ntl_ja.png
"""

import os
import math
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyArrowPatch

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)


def draw_ntl(lang="en"):
    if lang == "en":
        labels = [
            ("N", "Nosological\nAct"),
            ("V", "Patient\nValidation"),
            ("A", "Clinical\nAttention"),
            ("R", "Research\nInvestment"),
            ("T", "Treatment\nDevelopment"),
            ("N\u2032", "Nosological\nRefinement"),
        ]
        descs = [
            "Disease category\ncreated / refined",
            "Name reduces\ndiagnostic uncertainty",
            "New cognitive categories\nfor clinicians",
            "Funding & trials organized\naround category",
            "New therapeutics for\nnamed condition",
            "Category revised\nbased on evidence",
        ]
        center_title = "BOUNDARY\nCONDITION"
        center_desc = "Without empirical\ngrounding:\nDisease Mongering"
        fig_title = "Figure 1. The Nosological Therapeutic Loop (NTL)"
    else:
        labels = [
            ("N", "\u75be\u60a3\u547d\u540d\n\u884c\u70ba"),
            ("V", "\u60a3\u8005\u306e\n\u627f\u8a8d"),
            ("A", "\u81e8\u5e8a\u7684\n\u6ce8\u76ee"),
            ("R", "\u7814\u7a76\n\u6295\u8cc7"),
            ("T", "\u6cbb\u7642\n\u958b\u767a"),
            ("N\u2032", "\u75be\u60a3\u547d\u540d\u306e\n\u7cbe\u7dfb\u5316"),
        ]
        descs = [
            "\u75be\u60a3\u30ab\u30c6\u30b4\u30ea\u30fc\u306e\n\u5275\u8a2d\u30fb\u7cbe\u7dfb\u5316",
            "\u8a3a\u65ad\u7684\u4e0d\u78ba\u5b9f\u6027\u306e\n\u8efd\u6e1b",
            "\u65b0\u305f\u306a\u8a8d\u77e5\n\u30ab\u30c6\u30b4\u30ea\u30fc\u306e\u7372\u5f97",
            "\u8cc7\u91d1\u30fb\u8a66\u9a13\u306e\n\u30ab\u30c6\u30b4\u30ea\u30fc\u5468\u8fba\u3078\u306e\u7d44\u7e54\u5316",
            "\u547d\u540d\u3055\u308c\u305f\u72b6\u614b\u3078\u306e\n\u65b0\u6cbb\u7642\u6cd5",
            "\u30a8\u30d3\u30c7\u30f3\u30b9\u306b\u57fa\u3065\u304f\n\u30ab\u30c6\u30b4\u30ea\u30fc\u6539\u8a02",
        ]
        center_title = "\u5883\u754c\u6761\u4ef6"
        center_desc = "\u5b9f\u8a3c\u7684\u6839\u62e0\u306a\u304d\u5834\u5408:\n\u75be\u60a3\u5587\u4f1d"
        fig_title = "\u56f31. \u75be\u60a3\u547d\u540d\u6cbb\u7642\u30eb\u30fc\u30d7\uff08NTL\uff09"

    colors = ['#2962FF', '#009688', '#4CAF50', '#FF9800', '#E91E63', '#673AB7']

    if lang == "ja":
        import matplotlib.font_manager as fm
        jp_font_path = '/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf'
        jp_prop = fm.FontProperties(fname=jp_font_path)
        plt.rcParams['font.family'] = jp_prop.get_name()
        fm.fontManager.addfont(jp_font_path)
    else:
        plt.rcParams['font.family'] = 'DejaVu Sans'

    fig, ax = plt.subplots(1, 1, figsize=(10, 8))
    ax.set_xlim(-5.5, 5.5)
    ax.set_ylim(-5.5, 5.5)
    ax.set_aspect('equal')
    ax.axis('off')

    n = 6
    radius = 3.2
    node_radius = 0.85
    positions = []
    for i in range(n):
        angle = math.pi / 2 - 2 * math.pi * i / n
        x = radius * math.cos(angle)
        y = radius * math.sin(angle)
        positions.append((x, y))

    # Draw arrows
    for i in range(n):
        x1, y1 = positions[i]
        x2, y2 = positions[(i + 1) % n]
        dx, dy = x2 - x1, y2 - y1
        dist = math.sqrt(dx**2 + dy**2)
        # Shorten to avoid overlapping nodes
        shrink = node_radius + 0.15
        ratio = shrink / dist
        ax1 = x1 + dx * ratio
        ay1 = y1 + dy * ratio
        ax2 = x2 - dx * ratio
        ay2 = y2 - dy * ratio
        arrow = FancyArrowPatch(
            (ax1, ay1), (ax2, ay2),
            arrowstyle='->', mutation_scale=20,
            color='#888888', linewidth=2,
            connectionstyle='arc3,rad=0.08'
        )
        ax.add_patch(arrow)

    # Draw nodes
    for i, ((abbrev, label), desc, color) in enumerate(zip(labels, descs, colors)):
        x, y = positions[i]
        circle = plt.Circle((x, y), node_radius, color=color, ec='white', linewidth=2, zorder=3)
        ax.add_patch(circle)
        ax.text(x, y + 0.2, abbrev, ha='center', va='center',
                fontsize=16, fontweight='bold', color='white', zorder=4)
        ax.text(x, y - 0.25, label, ha='center', va='center',
                fontsize=8, color='white', zorder=4)

        # Description outside
        angle = math.pi / 2 - 2 * math.pi * i / n
        desc_r = radius + 1.6
        dx = desc_r * math.cos(angle)
        dy = desc_r * math.sin(angle)
        ax.text(dx, dy, desc, ha='center', va='center',
                fontsize=7.5, color='#555555',
                bbox=dict(boxstyle='round,pad=0.3', facecolor='#f5f5f5', edgecolor='#cccccc'))

    # Center boundary condition
    center_circle = plt.Circle((0, 0), 1.0, color='#F44336', ec='white', linewidth=2, zorder=3)
    ax.add_patch(center_circle)
    ax.text(0, 0.25, center_title, ha='center', va='center',
            fontsize=10, fontweight='bold', color='white', zorder=4)
    ax.text(0, -0.35, center_desc, ha='center', va='center',
            fontsize=7, color='#FFCDD2', zorder=4)

    ax.set_title(fig_title, fontsize=13, fontweight='bold', pad=15)

    plt.tight_layout()
    suffix = "en" if lang == "en" else "ja"
    output_path = os.path.join(OUTPUT_DIR, f"figure1_ntl_{suffix}.png")
    fig.savefig(output_path, dpi=300, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"Figure saved to: {output_path}")
    return output_path


if __name__ == "__main__":
    draw_ntl("en")
    draw_ntl("ja")
