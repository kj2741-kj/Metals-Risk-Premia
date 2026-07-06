"""
All charts for Paper 1.

Figure 2  – NGL-petrochemical processing spreads over time with regime shading
Figure 3  – Intra-spread equity curves (best params, all 5 spreads)
Figure 4  – Butane-Ethylene: equity curve + Sharpe heatmap
Figure 5  – Inter-spread equity curves (best params, all 5 pairs)
Figure 6  – Ethane-Ethylene vs Propane-Propylene: equity curve + Sharpe heatmap
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import seaborn as sns
from pathlib import Path

from config import REGIMES, REGIME_COLORS, CHARTS_DIR


# ── colour palette (copper / charcoal) ────────────────────────────────────────
LINE_COLORS = [
    "#B87333", "#5C4827", "#E8A87C", "#7D4E2D",
    "#A0522D", "#D2691E", "#8B4513",
]
EQUITY_COLORS = LINE_COLORS


def _ensure_dir() -> None:
    CHARTS_DIR.mkdir(parents=True, exist_ok=True)


def _shade_regimes(ax: plt.Axes) -> None:
    """Add regime background shading and a legend patch row."""
    patches = []
    for label, (start, end) in REGIMES.items():
        ax.axvspan(
            pd.Timestamp(start), pd.Timestamp(end),
            alpha=0.25, color=REGIME_COLORS[label], zorder=0,
        )
        patches.append(mpatches.Patch(color=REGIME_COLORS[label], alpha=0.5, label=label))
    return patches


# ── Figure 2 ──────────────────────────────────────────────────────────────────

def plot_spreads(spreads: pd.DataFrame, out_dir: Path | None = None) -> Path:
    """
    Figure 2: five intra-spread series over time.
    Returns path to saved PNG.
    """
    _ensure_dir()
    out_dir = out_dir or CHARTS_DIR

    fig, ax = plt.subplots(figsize=(14, 6))
    regime_patches = _shade_regimes(ax)

    for i, col in enumerate(spreads.columns):
        ax.plot(spreads.index, spreads[col],
                color=LINE_COLORS[i % len(LINE_COLORS)],
                linewidth=0.9, label=col)

    ax.axhline(0, color="#333333", linewidth=0.7, linestyle="--")
    ax.set_title("Figure 2 – NGL-Petrochemical Processing Spreads ($/gal)", fontsize=13)
    ax.set_ylabel("Spread ($/gal)")
    ax.set_xlabel("")

    line_handles, line_labels = ax.get_legend_handles_labels()
    first_legend = ax.legend(
        handles=line_handles,
        labels=line_labels,
        loc="upper left", fontsize=8, framealpha=0.8,
    )
    ax.add_artist(first_legend)
    ax.legend(handles=regime_patches, loc="upper right", fontsize=8, framealpha=0.8)

    fig.tight_layout()
    out_path = out_dir / "fig2_intra_spreads.png"
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    return out_path


# ── Figure 3 ──────────────────────────────────────────────────────────────────

def plot_intra_equity(
    pnl_dict: dict[str, pd.Series],
    out_dir: Path | None = None,
) -> Path:
    """
    Figure 3: cumulative P/L curves for all intra-spreads (best params).
    pnl_dict: {spread_name: daily_pnl Series}
    """
    _ensure_dir()
    out_dir = out_dir or CHARTS_DIR

    fig, ax = plt.subplots(figsize=(14, 6))
    regime_patches = _shade_regimes(ax)

    for i, (name, pnl) in enumerate(pnl_dict.items()):
        cum = pnl.fillna(0).cumsum()
        ax.plot(cum.index, cum.values,
                color=EQUITY_COLORS[i % len(EQUITY_COLORS)],
                linewidth=1.1, label=name)

    ax.axhline(0, color="#333333", linewidth=0.7, linestyle="--")
    ax.set_title("Figure 3 – Intra-Spread Strategies: Cumulative P/L ($/gal)", fontsize=13)
    ax.set_ylabel("Cumulative P/L ($/gal)")

    line_handles, line_labels = ax.get_legend_handles_labels()
    first_legend = ax.legend(handles=line_handles, labels=line_labels,
                             loc="upper left", fontsize=8, framealpha=0.8)
    ax.add_artist(first_legend)
    ax.legend(handles=regime_patches, loc="lower right", fontsize=8, framealpha=0.8)

    fig.tight_layout()
    out_path = out_dir / "fig3_intra_equity.png"
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    return out_path


# ── Figure 4 ──────────────────────────────────────────────────────────────────

def plot_intra_focus(
    spread_name: str,
    pnl: pd.Series,
    heatmap: pd.DataFrame,
    best_n: int,
    best_eps: float,
    out_dir: Path | None = None,
) -> Path:
    """
    Figure 4: equity curve (top) + Sharpe heatmap (bottom) for one intra-spread.
    Used primarily for Butane-Ethylene.
    """
    _ensure_dir()
    out_dir = out_dir or CHARTS_DIR

    fig, (ax_top, ax_bot) = plt.subplots(2, 1, figsize=(14, 10),
                                          gridspec_kw={"height_ratios": [1.4, 1]})

    # -- equity curve --
    regime_patches = _shade_regimes(ax_top)
    cum = pnl.fillna(0).cumsum()
    ax_top.plot(cum.index, cum.values, color="#B87333", linewidth=1.2,
                label=f"{spread_name} (n={best_n}, ε={best_eps:.2f})")
    ax_top.axhline(0, color="#333333", linewidth=0.7, linestyle="--")
    ax_top.set_title(
        f"Figure 4 – {spread_name}: Equity Curve & Sharpe Heatmap", fontsize=13)
    ax_top.set_ylabel("Cumulative P/L ($/gal)")

    line_handles, line_labels = ax_top.get_legend_handles_labels()
    first_legend = ax_top.legend(handles=line_handles, labels=line_labels,
                                  loc="upper left", fontsize=9, framealpha=0.8)
    ax_top.add_artist(first_legend)
    ax_top.legend(handles=regime_patches, loc="lower right", fontsize=8, framealpha=0.8)

    # -- heatmap --
    sns.heatmap(
        heatmap.astype(float),
        ax=ax_bot,
        cmap="RdYlGn",
        center=0,
        annot=True, fmt=".2f", annot_kws={"size": 7},
        linewidths=0.3,
        cbar_kws={"label": "Sharpe Ratio"},
    )
    ax_bot.set_title("Sharpe Ratio Grid (rows=Lookback, cols=Threshold)", fontsize=10)
    ax_bot.set_xlabel("Threshold ($/gal)")
    ax_bot.set_ylabel("Lookback Period (days)")

    # Mark best cell
    row_idx = list(heatmap.index).index(best_n)
    col_idx = list(heatmap.columns).index(best_eps)
    ax_bot.add_patch(plt.Rectangle(
        (col_idx, row_idx), 1, 1,
        fill=False, edgecolor="black", linewidth=2.5,
    ))

    fig.tight_layout()
    safe_name = spread_name.replace(" ", "_").replace("-", "_")
    out_path = out_dir / f"fig4_{safe_name}_focus.png"
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    return out_path


# ── Figure 5 ──────────────────────────────────────────────────────────────────

def plot_inter_equity(
    pnl_dict: dict[str, pd.Series],
    out_dir: Path | None = None,
) -> Path:
    """Figure 5: cumulative P/L curves for all inter-spreads (best params)."""
    _ensure_dir()
    out_dir = out_dir or CHARTS_DIR

    fig, ax = plt.subplots(figsize=(14, 6))
    regime_patches = _shade_regimes(ax)

    for i, (name, pnl) in enumerate(pnl_dict.items()):
        cum = pnl.fillna(0).cumsum()
        ax.plot(cum.index, cum.values,
                color=EQUITY_COLORS[i % len(EQUITY_COLORS)],
                linewidth=1.1, label=name)

    ax.axhline(0, color="#333333", linewidth=0.7, linestyle="--")
    ax.set_title("Figure 5 – Inter-Spread Strategies: Cumulative P/L ($/gal)", fontsize=13)
    ax.set_ylabel("Cumulative P/L ($/gal)")

    line_handles, line_labels = ax.get_legend_handles_labels()
    first_legend = ax.legend(handles=line_handles, labels=line_labels,
                             loc="upper left", fontsize=8, framealpha=0.8)
    ax.add_artist(first_legend)
    ax.legend(handles=regime_patches, loc="lower right", fontsize=8, framealpha=0.8)

    fig.tight_layout()
    out_path = out_dir / "fig5_inter_equity.png"
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    return out_path


# ── Figure 6 ──────────────────────────────────────────────────────────────────

def plot_inter_focus(
    inter_name: str,
    pnl: pd.Series,
    heatmap: pd.DataFrame,
    best_n: int,
    best_eps: float,
    out_dir: Path | None = None,
) -> Path:
    """
    Figure 6: equity curve (top) + Sharpe heatmap (bottom) for one inter-spread.
    Used primarily for Ethane-Ethylene vs Propane-Propylene.
    """
    _ensure_dir()
    out_dir = out_dir or CHARTS_DIR

    fig, (ax_top, ax_bot) = plt.subplots(2, 1, figsize=(14, 10),
                                          gridspec_kw={"height_ratios": [1.4, 1]})

    # -- equity curve --
    regime_patches = _shade_regimes(ax_top)
    cum = pnl.fillna(0).cumsum()
    ax_top.plot(cum.index, cum.values, color="#B87333", linewidth=1.2,
                label=f"{inter_name} (n={best_n}, ε={best_eps:.2f})")
    ax_top.axhline(0, color="#333333", linewidth=0.7, linestyle="--")
    ax_top.set_title(
        f"Figure 6 – {inter_name}: Equity Curve & Sharpe Heatmap", fontsize=13)
    ax_top.set_ylabel("Cumulative P/L ($/gal)")

    line_handles, line_labels = ax_top.get_legend_handles_labels()
    first_legend = ax_top.legend(handles=line_handles, labels=line_labels,
                                  loc="upper left", fontsize=9, framealpha=0.8)
    ax_top.add_artist(first_legend)
    ax_top.legend(handles=regime_patches, loc="lower right", fontsize=8, framealpha=0.8)

    # -- heatmap --
    sns.heatmap(
        heatmap.astype(float),
        ax=ax_bot,
        cmap="RdYlGn",
        center=0,
        annot=True, fmt=".2f", annot_kws={"size": 7},
        linewidths=0.3,
        cbar_kws={"label": "Sharpe Ratio"},
    )
    ax_bot.set_title("Sharpe Ratio Grid (rows=Lookback, cols=Threshold)", fontsize=10)
    ax_bot.set_xlabel("Threshold ($/gal)")
    ax_bot.set_ylabel("Lookback Period (days)")

    row_idx = list(heatmap.index).index(best_n)
    col_idx = list(heatmap.columns).index(best_eps)
    ax_bot.add_patch(plt.Rectangle(
        (col_idx, row_idx), 1, 1,
        fill=False, edgecolor="black", linewidth=2.5,
    ))

    fig.tight_layout()
    safe_name = inter_name.replace(" ", "_").replace("-", "_")
    out_path = out_dir / f"fig6_{safe_name}_focus.png"
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    return out_path


# ── Cointegration summary bar chart ───────────────────────────────────────────

def plot_coint_summary(
    intra_results: pd.DataFrame,
    inter_results: pd.DataFrame,
    out_dir: Path | None = None,
) -> Path:
    """Bar chart of p-values for all cointegration tests with 5% threshold line."""
    _ensure_dir()
    out_dir = out_dir or CHARTS_DIR

    combined = pd.concat(
        [intra_results[["P-Value"]], inter_results[["P-Value"]]],
        keys=["Intra", "Inter"],
    )
    combined = combined.reset_index(level=0).rename(columns={"level_0": "Type"})
    combined = combined.reset_index()

    fig, ax = plt.subplots(figsize=(14, 5))
    colors = ["#B87333" if t == "Intra" else "#5C4827" for t in combined["Type"]]
    bars = ax.bar(combined.iloc[:, 1], combined["P-Value"], color=colors)
    ax.axhline(0.05, color="red", linewidth=1.2, linestyle="--", label="5% significance")
    ax.set_ylabel("P-Value (Engle-Granger)")
    ax.set_title("Cointegration Test P-Values – Intra and Inter Spreads", fontsize=12)
    plt.xticks(rotation=30, ha="right", fontsize=7)
    ax.legend(fontsize=9)

    intra_patch = mpatches.Patch(color="#B87333", label="Intra-spread")
    inter_patch = mpatches.Patch(color="#5C4827", label="Inter-spread")
    ax.legend(handles=[intra_patch, inter_patch,
                        mpatches.Patch(color="red", label="5% level")],
              fontsize=9)

    fig.tight_layout()
    out_path = out_dir / "coint_pvalues.png"
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    return out_path
