"""
Charts for Paper 2.

Figure 1  – Roll-adjusted continuous price series for all 9 instruments
Figure 2  – Cumulative returns: Momentum (EW + DRP)
Figure 3  – Cumulative returns: Carry (EW + DRP)
Figure 4  – Cumulative returns: Value (EW + DRP)
Figure 5  – Cumulative returns: Stat-Arb spreads
Figure 6  – Combined portfolio cumulative returns vs each strategy
Figure 7  – Strategy correlation heatmap
Figure 8  – IR bar chart across strategies and sub-samples
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import seaborn as sns
from pathlib import Path

from config import CHARTS_DIR, SUBSAMPLE_1, SUBSAMPLE_2, INSTRUMENTS


LINE_COLORS = [
    "#B87333", "#5C4827", "#E8A87C", "#7D4E2D",
    "#A0522D", "#D2691E", "#8B4513", "#C08040", "#6B3A2A",
]

SCHEME_COLORS = {"EW": "#B87333", "DRP": "#5C4827"}


def _ensure_dir() -> None:
    CHARTS_DIR.mkdir(parents=True, exist_ok=True)


def _shade_subsamples(ax: plt.Axes) -> list:
    patches = []
    for label, (s, e), color in [
        ("Sub-sample 1 (2015–2022)", SUBSAMPLE_1, "#D4EDDA"),
        ("Sub-sample 2 (2022–2025)", SUBSAMPLE_2, "#CCE5FF"),
    ]:
        ax.axvspan(pd.Timestamp(s), pd.Timestamp(e),
                   alpha=0.2, color=color, zorder=0)
        patches.append(mpatches.Patch(color=color, alpha=0.5, label=label))
    return patches


# ── Figure 1 ──────────────────────────────────────────────────────────────────

def plot_continuous_prices(
    cont: dict[str, tuple[pd.Series, pd.Series]],
    out_dir: Path | None = None,
) -> Path:
    """Figure 1: roll-adjusted prices for all instruments ($/bbl equivalent)."""
    _ensure_dir()
    out_dir = out_dir or CHARTS_DIR

    n = len(cont)
    fig, axes = plt.subplots(
        (n + 2) // 3, 3,
        figsize=(18, 4 * ((n + 2) // 3)),
        sharex=False,
    )
    axes_flat = axes.flatten()

    for i, (ticker, (adj, _)) in enumerate(cont.items()):
        ax = axes_flat[i]
        ax.plot(adj.index, adj.values, color=LINE_COLORS[i % len(LINE_COLORS)],
                linewidth=0.8)
        name = INSTRUMENTS[ticker]["name"]
        ax.set_title(f"{ticker} – {name}", fontsize=9)
        ax.set_ylabel("$/bbl equiv.", fontsize=7)
        ax.tick_params(axis="both", labelsize=7)

    for j in range(i + 1, len(axes_flat)):
        axes_flat[j].set_visible(False)

    fig.suptitle("Figure 1 – Roll-Adjusted Continuous Price Series ($/bbl)", fontsize=12, y=1.01)
    fig.tight_layout()
    out_path = out_dir / "fig1_continuous_prices.png"
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return out_path


# ── Helper: single-strategy cumulative return plot ────────────────────────────

def _plot_strategy_cum(
    ew_ret: pd.Series,
    drp_ret: pd.Series,
    title: str,
    out_path: Path,
) -> None:
    fig, ax = plt.subplots(figsize=(14, 5))
    shade_patches = _shade_subsamples(ax)

    ax.plot(ew_ret.cumsum().index,  ew_ret.cumsum().values,
            color=SCHEME_COLORS["EW"],  linewidth=1.2, label="Equal Weight (EW)")
    ax.plot(drp_ret.cumsum().index, drp_ret.cumsum().values,
            color=SCHEME_COLORS["DRP"], linewidth=1.2, label="Dynamic Risk Parity (DRP)", linestyle="--")
    ax.axhline(0, color="#333", linewidth=0.6, linestyle=":")
    ax.set_title(title, fontsize=12)
    ax.set_ylabel("Cumulative Log-Return")

    h1, l1 = ax.get_legend_handles_labels()
    leg1 = ax.legend(handles=h1, labels=l1, loc="upper left", fontsize=9, framealpha=0.8)
    ax.add_artist(leg1)
    ax.legend(handles=shade_patches, loc="lower right", fontsize=8, framealpha=0.8)

    fig.tight_layout()
    fig.savefig(out_path, dpi=150)
    plt.close(fig)


# ── Figures 2-5 ───────────────────────────────────────────────────────────────

def plot_momentum_equity(ew: pd.Series, drp: pd.Series, out_dir: Path | None = None) -> Path:
    _ensure_dir(); out_dir = out_dir or CHARTS_DIR
    p = out_dir / "fig2_momentum_equity.png"
    _plot_strategy_cum(ew, drp, "Figure 2 – Momentum Strategy Cumulative Returns", p)
    return p


def plot_carry_equity(ew: pd.Series, drp: pd.Series, out_dir: Path | None = None) -> Path:
    _ensure_dir(); out_dir = out_dir or CHARTS_DIR
    p = out_dir / "fig3_carry_equity.png"
    _plot_strategy_cum(ew, drp, "Figure 3 – Carry Strategy Cumulative Returns", p)
    return p


def plot_value_equity(ew: pd.Series, drp: pd.Series, out_dir: Path | None = None) -> Path:
    _ensure_dir(); out_dir = out_dir or CHARTS_DIR
    p = out_dir / "fig4_value_equity.png"
    _plot_strategy_cum(ew, drp, "Figure 4 – Value Strategy Cumulative Returns", p)
    return p


def plot_statarb_equity(
    spread_returns: dict[str, pd.Series],
    out_dir: Path | None = None,
) -> Path:
    _ensure_dir(); out_dir = out_dir or CHARTS_DIR

    fig, ax = plt.subplots(figsize=(14, 5))
    shade_patches = _shade_subsamples(ax)

    for i, (name, ret) in enumerate(spread_returns.items()):
        ax.plot(ret.cumsum().index, ret.cumsum().values,
                color=LINE_COLORS[i % len(LINE_COLORS)], linewidth=0.9, label=name)

    ax.axhline(0, color="#333", linewidth=0.6, linestyle=":")
    ax.set_title("Figure 5 – Stat-Arb Spread Cumulative Returns ($/bbl)", fontsize=12)
    ax.set_ylabel("Cumulative Log-Return")

    h1, l1 = ax.get_legend_handles_labels()
    leg1 = ax.legend(handles=h1, labels=l1, loc="upper left", fontsize=7, framealpha=0.8, ncol=2)
    ax.add_artist(leg1)
    ax.legend(handles=shade_patches, loc="lower right", fontsize=8, framealpha=0.8)

    fig.tight_layout()
    p = out_dir / "fig5_statarb_equity.png"
    fig.savefig(p, dpi=150)
    plt.close(fig)
    return p


# ── Figure 6 ──────────────────────────────────────────────────────────────────

def plot_combined_equity(
    strategy_rets: dict[str, pd.Series],
    combined: pd.Series,
    out_dir: Path | None = None,
) -> Path:
    _ensure_dir(); out_dir = out_dir or CHARTS_DIR

    fig, ax = plt.subplots(figsize=(14, 5))
    shade_patches = _shade_subsamples(ax)

    for i, (name, ret) in enumerate(strategy_rets.items()):
        ax.plot(ret.cumsum().index, ret.cumsum().values,
                color=LINE_COLORS[i % len(LINE_COLORS)],
                linewidth=0.8, alpha=0.7, label=name)

    ax.plot(combined.cumsum().index, combined.cumsum().values,
            color="black", linewidth=2.0, label="Combined Portfolio", zorder=5)
    ax.axhline(0, color="#333", linewidth=0.6, linestyle=":")
    ax.set_title("Figure 6 – Combined Portfolio vs Individual Strategies", fontsize=12)
    ax.set_ylabel("Cumulative Log-Return")

    h1, l1 = ax.get_legend_handles_labels()
    leg1 = ax.legend(handles=h1, labels=l1, loc="upper left", fontsize=8, framealpha=0.8)
    ax.add_artist(leg1)
    ax.legend(handles=shade_patches, loc="lower right", fontsize=8, framealpha=0.8)

    fig.tight_layout()
    p = out_dir / "fig6_combined_equity.png"
    fig.savefig(p, dpi=150)
    plt.close(fig)
    return p


# ── Figure 7 ──────────────────────────────────────────────────────────────────

def plot_correlation_heatmap(
    strategy_rets: dict[str, pd.Series],
    out_dir: Path | None = None,
) -> Path:
    _ensure_dir(); out_dir = out_dir or CHARTS_DIR

    df   = pd.DataFrame(strategy_rets).dropna(how="all")
    corr = df.corr()

    fig, ax = plt.subplots(figsize=(10, 8))
    sns.heatmap(corr, ax=ax, cmap="RdYlGn", center=0,
                annot=True, fmt=".2f", annot_kws={"size": 8},
                linewidths=0.4, vmin=-1, vmax=1)
    ax.set_title("Figure 7 – Strategy Return Correlations", fontsize=12)
    fig.tight_layout()
    p = out_dir / "fig7_correlation_heatmap.png"
    fig.savefig(p, dpi=150)
    plt.close(fig)
    return p


# ── Figure 8 ──────────────────────────────────────────────────────────────────

def plot_ir_bars(
    perf_df: pd.DataFrame,
    out_dir: Path | None = None,
) -> Path:
    """
    Figure 8: IR bar chart from the combined performance DataFrame.
    perf_df must have 'IR' column and a descriptive index.
    """
    _ensure_dir(); out_dir = out_dir or CHARTS_DIR

    ir = perf_df["IR"].dropna()
    colors = ["#B87333" if v >= 0 else "#8B1A1A" for v in ir.values]

    fig, ax = plt.subplots(figsize=(14, 5))
    ax.bar(range(len(ir)), ir.values, color=colors, edgecolor="white", linewidth=0.5)
    ax.set_xticks(range(len(ir)))
    ax.set_xticklabels(ir.index, rotation=30, ha="right", fontsize=7)
    ax.axhline(0, color="black", linewidth=0.8)
    ax.set_ylabel("Information Ratio")
    ax.set_title("Figure 8 – Information Ratios by Strategy and Period", fontsize=12)
    fig.tight_layout()
    p = out_dir / "fig8_ir_bars.png"
    fig.savefig(p, dpi=150)
    plt.close(fig)
    return p
