# -*- coding: utf-8 -*-
"""Standalone Figure B-2 and D-1 PNGs for Reflective Report (does not modify the docx)."""
from __future__ import annotations

import os
import shutil
import sys

OUT_DIR = (
    r"D:\桌面文档文件勿动！\Desktop\毕业设计\论文\个人部分2\毕业论文全套_生成"
)


def generate_figure_b2(out_dir: str) -> str:
    """Compact 3+2 layout (~square). Both rows centered in x∈[0,10] so nothing clips right."""
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

    os.makedirs(out_dir, exist_ok=True)

    fig, ax = plt.subplots(figsize=(6.6, 6.4))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.axis("off")

    layers = [
        ("Presentation", "HTML / JS / CSS\n/static"),
        ("API", "FastAPI routers\nOAuth2 JWT"),
        ("Domain", "Sessions, emotions,\nquestions, recordings"),
        ("Persistence", "SQLite + ORM\nmodels"),
        ("ML", "Decode → quality\nDeepFace emotion"),
    ]

    gap = 0.14
    box_h = 1.02
    y_top = 6.55
    y_bot = 3.45
    cx = 5.0
    row1_outer = 9.2
    box_w = (row1_outer - 2 * gap) / 3
    x0 = cx - row1_outer / 2

    row2_outer = 2 * box_w + gap
    x2_0 = cx - row2_outer / 2

    def place(left: float, bottom: float, title: str, sub: str) -> None:
        ax.add_patch(
            FancyBboxPatch(
                (left, bottom),
                box_w,
                box_h,
                boxstyle="round,pad=0.03",
                ec="#333",
                fc="#f0f4f8",
            )
        )
        ax.text(left + box_w / 2, bottom + box_h * 0.68, title, ha="center", va="center", fontsize=8.0, fontweight="bold")
        ax.text(left + box_w / 2, bottom + box_h * 0.33, sub, ha="center", va="center", fontsize=6.9)

    def arrow(x1: float, y1: float, x2: float, y2: float) -> None:
        ax.add_patch(
            FancyArrowPatch(
                (x1, y1),
                (x2, y2),
                arrowstyle="->",
                mutation_scale=11,
                color="#555",
                linewidth=1.0,
            )
        )

    for i in range(3):
        place(x0 + i * (box_w + gap), y_top, layers[i][0], layers[i][1])

    place(x2_0, y_bot, layers[3][0], layers[3][1])
    place(x2_0 + box_w + gap, y_bot, layers[4][0], layers[4][1])

    cy_top = y_top + box_h / 2
    for i in range(2):
        xf = x0 + i * (box_w + gap) + box_w
        xt = x0 + (i + 1) * (box_w + gap)
        arrow(xf, cy_top, xt, cy_top)

    x_domain_left = x0 + 2 * (box_w + gap)
    xc_d = x_domain_left + box_w / 2
    xc_p = x2_0 + box_w / 2
    y_mid = (y_top + y_bot + box_h) / 2
    arrow(xc_d, y_top, xc_d, y_mid)
    arrow(xc_d, y_mid, xc_p, y_mid)
    arrow(xc_p, y_mid, xc_p, y_bot + box_h)

    cy_bot = y_bot + box_h / 2
    arrow(x2_0 + box_w, cy_bot, x2_0 + box_w + gap, cy_bot)

    fig.suptitle(
        "Figure B-2. Layered structure (Web1 delivery view).",
        fontsize=11,
        fontweight="bold",
        y=0.96,
    )
    out_b2 = os.path.join(out_dir, "Figure_B-2_Layered_structure.png")
    fig.savefig(out_b2, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return out_b2


def generate_figure_d1(out_dir: str) -> str:
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

    os.makedirs(out_dir, exist_ok=True)

    fig, ax = plt.subplots(figsize=(6.8, 6.2))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.axis("off")

    steps = [
        "WS JSON\n(base64)",
        "imdecode",
        "recognize\nHaar→ROI",
        "DeepFace\nanalyze",
        "smooth\n(5-frame)",
        "DB batch\n/3 frames",
        "JPEG +\nWS reply",
    ]
    box_w, box_h = 1.95, 1.05
    gap = 0.22
    y_top = 6.35
    y_bot = 3.15

    def place_box(left: float, bottom: float, label: str) -> None:
        ax.add_patch(
            FancyBboxPatch(
                (left, bottom),
                box_w,
                box_h,
                boxstyle="round,pad=0.03",
                ec="#444",
                fc="#fff8e6",
            )
        )
        ax.text(left + box_w / 2, bottom + box_h / 2, label, ha="center", va="center", fontsize=7.5)

    def arrow_seg(x1: float, y1: float, x2: float, y2: float) -> None:
        ax.add_patch(
            FancyArrowPatch(
                (x1, y1),
                (x2, y2),
                arrowstyle="->",
                mutation_scale=12,
                color="#666",
                linewidth=1.0,
            )
        )

    row1_n, row2_n = 4, 3
    row1_w = row1_n * box_w + (row1_n - 1) * gap
    row2_w = row2_n * box_w + (row2_n - 1) * gap
    x1_0 = (10 - row1_w) / 2
    x2_0 = (10 - row2_w) / 2

    for i in range(row1_n):
        place_box(x1_0 + i * (box_w + gap), y_top, steps[i])
    for i in range(row2_n):
        place_box(x2_0 + i * (box_w + gap), y_bot, steps[row1_n + i])

    cy_top = y_top + box_h / 2
    for i in range(row1_n - 1):
        x_from = x1_0 + i * (box_w + gap) + box_w
        x_to = x1_0 + (i + 1) * (box_w + gap)
        arrow_seg(x_from, cy_top, x_to, cy_top)

    x_deep_r = x1_0 + 3 * (box_w + gap) + box_w
    x_smooth_l = x2_0
    mid_y = y_top - 0.35
    arrow_seg(x_deep_r, cy_top, x_deep_r, mid_y)
    arrow_seg(x_deep_r, mid_y, x_smooth_l, mid_y)
    arrow_seg(x_smooth_l, mid_y, x_smooth_l + box_w / 2, y_bot + box_h)

    cy_bot = y_bot + box_h / 2
    for i in range(row2_n - 1):
        x_from = x2_0 + i * (box_w + gap) + box_w
        x_to = x2_0 + (i + 1) * (box_w + gap)
        arrow_seg(x_from, cy_bot, x_to, cy_bot)

    fig.suptitle(
        "Figure D-1. Real-time pipeline (/ws/{session_id}).",
        fontsize=11,
        fontweight="bold",
        y=0.96,
    )
    out_d1 = os.path.join(out_dir, "Figure_D-1_Real-time_pipeline.png")
    fig.savefig(out_d1, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return out_d1


def _copy_to_assets(*paths: str) -> None:
    _here = os.path.dirname(os.path.abspath(__file__))
    _assets = os.path.join(_here, "assets")
    os.makedirs(_assets, exist_ok=True)
    for p in paths:
        if os.path.isfile(p):
            shutil.copy2(p, os.path.join(_assets, os.path.basename(p)))


def main_b2_only() -> None:
    out_b2 = generate_figure_b2(OUT_DIR)
    _copy_to_assets(out_b2)
    print(out_b2)


def main() -> None:
    out_b2 = generate_figure_b2(OUT_DIR)
    out_d1 = generate_figure_d1(OUT_DIR)
    _copy_to_assets(out_b2, out_d1)
    print(out_b2)
    print(out_d1)


if __name__ == "__main__":
    if "--b2-only" in sys.argv:
        main_b2_only()
    else:
        main()
