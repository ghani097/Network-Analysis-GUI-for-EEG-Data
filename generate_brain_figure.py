"""
Generate publication-quality brain surface figures showing the 28 ROIs
of the triple network (DMN, CEN, SN) mapped via the Destrieux atlas.

Produces lateral + medial views for both hemispheres.

Requirements:
    pip install nilearn matplotlib numpy

Usage:
    python generate_brain_figure.py
"""

import warnings
warnings.filterwarnings("ignore")

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.colors import ListedColormap
from matplotlib.patches import Patch
from nilearn import datasets, plotting

# --------------------------------------------------------------------------
# Mapping: Desikan-Killiany ROIs -> Destrieux atlas labels
# --------------------------------------------------------------------------
# The Destrieux atlas (76 labels) is a finer parcellation than DK.
# Each DK region maps to one or more Destrieux labels.
# We group by network assignment (DMN, CEN, SN).

# Network code: 0=background, 1=DMN, 2=CEN, 3=SN
NETWORK_CODE = {"DMN": 1, "CEN": 2, "SN": 3}

# Destrieux label indices for each network
# (Indices match the order from nilearn.datasets.fetch_atlas_surf_destrieux)
#
# DMN regions (16 ROIs = 8 bilateral):
#   inferiorparietal      -> G_pariet_inf-Angular(25), G_pariet_inf-Supramar(26)
#   isthmuscingulate      -> G_cingul-Post-ventral(10), S_subparietal(72)
#   medialorbitofrontal   -> G_rectus(31), S_orbital_med-olfact(64), G_subcallosal(32)
#   middletemporal        -> G_temporal_middle(38)
#   parahippocampal       -> G_oc-temp_med-Parahip(23)
#   posteriorcingulate    -> G_cingul-Post-dorsal(9)
#   precuneus             -> G_precuneus(30)
#   rostralanteriorcingulate -> G_and_S_cingul-Ant(6)
#
# CEN regions (8 ROIs = 4 bilateral):
#   caudalmiddlefrontal   -> G_front_middle(15) [posterior part - shared with rostral]
#   rostralmiddlefrontal  -> G_front_middle(15), S_front_middle(54)
#   superiorfrontal       -> G_front_sup(16), S_front_sup(55)
#   superiorparietal      -> G_parietal_sup(27), S_intrapariet_and_P_trans(57)
#
# SN regions (4 ROIs = 2 bilateral):
#   caudalanteriorcingulate -> G_and_S_cingul-Mid-Ant(7)
#   insula                  -> G_insular_short(18), G_Ins_lg_and_S_cent_ins(17),
#                              S_circular_insula_ant(48), S_circular_insula_inf(49),
#                              S_circular_insula_sup(50)

DESTRIEUX_TO_NETWORK = {
    # DMN
    25: 1,  # G_pariet_inf-Angular
    26: 1,  # G_pariet_inf-Supramar
    10: 1,  # G_cingul-Post-ventral (isthmus cingulate)
    72: 1,  # S_subparietal (isthmus cingulate)
    31: 1,  # G_rectus (medial orbitofrontal)
    64: 1,  # S_orbital_med-olfact (medial orbitofrontal)
    32: 1,  # G_subcallosal (medial orbitofrontal)
    38: 1,  # G_temporal_middle
    23: 1,  # G_oc-temp_med-Parahip (parahippocampal)
    9:  1,  # G_cingul-Post-dorsal (posterior cingulate)
    30: 1,  # G_precuneus
    6:  1,  # G_and_S_cingul-Ant (rostral anterior cingulate)

    # CEN
    15: 2,  # G_front_middle (caudal + rostral middle frontal)
    54: 2,  # S_front_middle
    16: 2,  # G_front_sup (superior frontal)
    55: 2,  # S_front_sup
    27: 2,  # G_parietal_sup (superior parietal)
    57: 2,  # S_intrapariet_and_P_trans

    # SN
    7:  3,  # G_and_S_cingul-Mid-Ant (caudal anterior cingulate)
    18: 3,  # G_insular_short
    17: 3,  # G_Ins_lg_and_S_cent_ins
    48: 3,  # S_circular_insula_ant
    49: 3,  # S_circular_insula_inf
    50: 3,  # S_circular_insula_sup
}

# Colours matching your original BrainStorm figure
NETWORK_COLORS = {
    "DMN": [0.0, 0.8, 0.5, 1.0],   # Green
    "CEN": [0.6, 0.3, 0.8, 1.0],   # Purple
    "SN":  [0.9, 0.2, 0.3, 1.0],   # Red
}


def build_roi_map(label_map):
    """Convert Destrieux per-vertex labels to network codes."""
    roi_map = np.zeros(len(label_map), dtype=float)
    for destrieux_idx, network_code in DESTRIEUX_TO_NETWORK.items():
        roi_map[label_map == destrieux_idx] = network_code
    return roi_map


def make_cmap():
    """Custom colormap: 0=grey bg, 1=DMN green, 2=CEN purple, 3=SN red."""
    return ListedColormap([
        [0.78, 0.78, 0.78, 1.0],
        NETWORK_COLORS["DMN"],
        NETWORK_COLORS["CEN"],
        NETWORK_COLORS["SN"],
    ])


def make_legend():
    return [
        Patch(facecolor=NETWORK_COLORS["CEN"], label="CEN (Central Executive) - 8 ROIs"),
        Patch(facecolor=NETWORK_COLORS["DMN"], label="DMN (Default Mode) - 16 ROIs"),
        Patch(facecolor=NETWORK_COLORS["SN"],  label="SN (Salience) - 4 ROIs"),
    ]


def generate_brain_figure(output_path="brain_triple_network.png", dpi=300):
    """Generate a 2x2 figure: lateral + medial views for both hemispheres."""

    fsaverage = datasets.fetch_surf_fsaverage(mesh="fsaverage5")
    destrieux = datasets.fetch_atlas_surf_destrieux()

    lh_roi_map = build_roi_map(destrieux.map_left)
    rh_roi_map = build_roi_map(destrieux.map_right)

    # Report coverage
    for name, rmap in [("Left", lh_roi_map), ("Right", rh_roi_map)]:
        total = len(rmap)
        colored = np.sum(rmap > 0)
        print(f"{name} hemisphere: {colored}/{total} vertices colored "
              f"({100*colored/total:.1f}%)")

    custom_cmap = make_cmap()

    # Generate 4 separate images with no titles, labels, or legends
    view_configs = [
        ("left",  "lateral", lh_roi_map, "left",  "left_lateral"),
        ("right", "lateral", rh_roi_map, "right", "right_lateral"),
        ("left",  "medial",  lh_roi_map, "left",  "left_medial"),
        ("right", "medial",  rh_roi_map, "right", "right_medial"),
        ("left",  "dorsal",  lh_roi_map, "left",  "left_dorsal"),
        ("right", "dorsal",  rh_roi_map, "right", "right_dorsal"),
    ]

    base = output_path.replace(".png", "")
    for hemi, view, data, mesh_side, name in view_configs:
        fig, ax = plt.subplots(1, 1, figsize=(6, 5),
                                subplot_kw={"projection": "3d"})
        plotting.plot_surf_roi(
            fsaverage[f"pial_{mesh_side}"],
            roi_map=data,
            hemi=hemi,
            view=view,
            bg_map=fsaverage[f"sulc_{mesh_side}"],
            bg_on_data=True,
            cmap=custom_cmap,
            vmin=0,
            vmax=3,
            axes=ax,
            figure=fig,
            alpha=0.85,
            colorbar=False,
        )
        fpath = f"{base}_{name}.png"
        fig.savefig(fpath, dpi=dpi, bbox_inches="tight",
                    facecolor="white", edgecolor="none", pad_inches=0.01)
        print(f"  Saved: {fpath}")
        plt.close(fig)

    print(f"\n4 separate images saved with prefix: {base}_")


def generate_all_views(output_path="brain_triple_network_all_views.png", dpi=300):
    """Generate lateral + medial + dorsal views (3x2 = 6 panels)."""

    fsaverage = datasets.fetch_surf_fsaverage(mesh="fsaverage5")
    destrieux = datasets.fetch_atlas_surf_destrieux()

    lh_roi_map = build_roi_map(destrieux.map_left)
    rh_roi_map = build_roi_map(destrieux.map_right)

    custom_cmap = make_cmap()

    fig = plt.figure(figsize=(16, 14), facecolor="white")

    configs = [
        (1, "left",  "lateral", lh_roi_map, "left",  "Left Lateral"),
        (2, "right", "lateral", rh_roi_map, "right", "Right Lateral"),
        (3, "left",  "medial",  lh_roi_map, "left",  "Left Medial"),
        (4, "right", "medial",  rh_roi_map, "right", "Right Medial"),
        (5, "left",  "dorsal",  lh_roi_map, "left",  "Left Dorsal"),
        (6, "right", "dorsal",  rh_roi_map, "right", "Right Dorsal"),
    ]

    for pos, hemi, view, data, mesh_side, title in configs:
        ax = fig.add_subplot(3, 2, pos, projection="3d")
        plotting.plot_surf_roi(
            fsaverage[f"pial_{mesh_side}"],
            roi_map=data,
            hemi=hemi,
            view=view,
            bg_map=fsaverage[f"sulc_{mesh_side}"],
            bg_on_data=True,
            cmap=custom_cmap,
            vmin=0,
            vmax=3,
            axes=ax,
            figure=fig,
            alpha=0.85,
            colorbar=False,
        )
        ax.set_title(title, fontsize=12, fontweight="bold", pad=8)

    fig.suptitle(
        "ROIs Forming a Triple Network (DMN, CEN, and SN)\n"
        "Lateral, Medial, and Dorsal Views - 28 ROIs (Desikan-Killiany Atlas)",
        fontsize=14, fontweight="bold", y=0.99
    )

    fig.legend(
        handles=make_legend(),
        loc="lower center",
        ncol=3,
        fontsize=11,
        frameon=True,
        bbox_to_anchor=(0.5, 0.005),
    )

    plt.tight_layout(rect=[0, 0.04, 1, 0.96])
    fig.savefig(output_path, dpi=dpi, bbox_inches="tight",
                facecolor="white", edgecolor="none")
    print(f"All-views figure saved to: {output_path}")
    plt.close(fig)


if __name__ == "__main__":
    print("=" * 60)
    print("Generating Triple Network Brain Figures")
    print("=" * 60)

    # Figure 1: Lateral + Medial views (4 panels)
    generate_brain_figure("brain_triple_network.png", dpi=300)

    # Figure 2: Lateral + Medial + Dorsal views (6 panels)
    generate_all_views("brain_triple_network_all_views.png", dpi=300)

    print("\nDone! Check the output PNG files.")
