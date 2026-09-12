from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np


RANDOM_SEED = 42
NUM_KERNELS = 36

OUTPUT_PATH = Path("images/example/wasserstein_example.png")

COLOR_OPTION_A = "tab:orange"
COLOR_REFERENCE_Q = "tab:blue"
COLOR_OPTION_B = "tab:red"


def wasserstein_1d_equal_weights(values_a, values_b):
    """
    Compute W1 for equally weighted one-dimensional empirical
    distributions containing the same number of support points.
    """
    sorted_a = np.sort(values_a)
    sorted_b = np.sort(values_b)

    return np.mean(np.abs(sorted_a - sorted_b))


def make_compact_supports(center, num_supports, spread):
    """
    Create compact supports with greater density close to the centre.
    """
    quantiles = np.linspace(0.04, 0.96, num_supports)

    supports = np.tan(np.pi * (quantiles - 0.5))
    supports = np.clip(supports, -2.0, 2.0)
    supports = supports / np.max(np.abs(supports))

    return center + spread * supports


def make_wide_supports(center, num_supports, spread):
    """
    Create broad supports that occupy a much wider horizontal interval.
    """
    quantiles = np.linspace(0.02, 0.98, num_supports)

    supports = np.tan(np.pi * (quantiles - 0.5))
    supports = np.clip(supports, -5.0, 5.0)
    supports = supports / np.max(np.abs(supports))

    return center + spread * supports


def plot_support_row(
    axis,
    x_values,
    y_position,
    color,
    rng,
    point_size=115,
    jitter_std=0.035,
):
    """Plot a horizontal kernel-support row with small vertical jitter."""
    y_values = y_position + rng.normal(
        loc=0.0,
        scale=jitter_std,
        size=len(x_values),
    )

    axis.scatter(
        x_values,
        y_values,
        s=point_size,
        color=color,
        alpha=0.82,
        edgecolors="none",
        zorder=3,
    )


def main():
    rng = np.random.default_rng(RANDOM_SEED)

    # Reference Q: compact distribution around x = 2.0.
    reference_q = make_compact_supports(
        center=2.0,
        num_supports=NUM_KERNELS,
        spread=1.25,
    )

    # Option A: compact and close to Q.
    option_a = make_compact_supports(
        center=2.30,
        num_supports=NUM_KERNELS,
        spread=1.32,
    )

    # Option B: wider and farther away from Q.
    option_b = make_wide_supports(
        center=-0.80,
        num_supports=NUM_KERNELS,
        spread=5.30,
    )

    # Calculate Wasserstein distances before plotting.
    distance_a_to_q = wasserstein_1d_equal_weights(
        option_a,
        reference_q,
    )

    distance_b_to_q = wasserstein_1d_equal_weights(
        option_b,
        reference_q,
    )

    OUTPUT_PATH.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    figure, axis = plt.subplots(
        figsize=(12.5, 4.8),
        layout="constrained",
        facecolor="white",
    )

    row_a = 2
    row_q = 1
    row_b = 0

    for row_position in [row_a, row_q, row_b]:
        axis.axhline(
            row_position,
            color="#D7D7D7",
            linewidth=1.2,
            zorder=1,
        )

    plot_support_row(
        axis=axis,
        x_values=option_a,
        y_position=row_a,
        color=COLOR_OPTION_A,
        rng=rng,
    )

    plot_support_row(
        axis=axis,
        x_values=reference_q,
        y_position=row_q,
        color=COLOR_REFERENCE_Q,
        rng=rng,
    )

    plot_support_row(
        axis=axis,
        x_values=option_b,
        y_position=row_b,
        color=COLOR_OPTION_B,
        rng=rng,
    )

    # Only the minimal A, Q, and B labels appear in the plot.
    axis.set_yticks(
        [row_a, row_q, row_b],
        ["A", "Q", "B"],
        fontsize=18,
        fontweight="bold",
    )

    axis.set_xlim(-7.5, 8.5)
    axis.set_ylim(-0.55, 2.55)

    axis.set_xlabel(
        "Kernel support location",
        fontsize=15,
        labelpad=10,
    )

    axis.set_ylabel("")

    axis.grid(
        axis="x",
        linestyle="--",
        linewidth=1.0,
        color="#DDDDDD",
        alpha=0.9,
    )

    axis.tick_params(
        axis="x",
        labelsize=12,
        length=5,
        width=1.0,
    )

    axis.tick_params(
        axis="y",
        length=0,
        pad=12,
    )

    axis.spines["top"].set_visible(False)
    axis.spines["right"].set_visible(False)
    axis.spines["left"].set_linewidth(1.2)
    axis.spines["bottom"].set_linewidth(1.2)

    figure.savefig(
        OUTPUT_PATH,
        dpi=180,
        bbox_inches="tight",
        pad_inches=0.08,
        facecolor="white",
    )

    plt.close(figure)

    print(f"Saved clean example plot to: {OUTPUT_PATH}")
    print(f"W(A, Q) = {distance_a_to_q:.4f}")
    print(f"W(B, Q) = {distance_b_to_q:.4f}")


if __name__ == "__main__":
    main()