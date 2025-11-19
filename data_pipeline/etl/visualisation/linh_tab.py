import os
import sys
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st
import plotly.express as px

# -----------------------------------------------------------------------------
# Streamlit global styling: Montserrat + dark background
# -----------------------------------------------------------------------------
st.markdown(
    """
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700&display=swap" rel="stylesheet">

    <style>
    /* Apply Montserrat everywhere in the main app */
    html, body, [class^="css"], [class*="css"],
    h1, h2, h3, h4, h5, h6,
    .stMarkdown, .stText, .stCaption,
    .stButton > button,
    div[data-testid="stMarkdownContainer"],
    div[data-testid="stHeader"] * {
        font-family: 'Montserrat', sans-serif !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# Path + DB connection import
# -----------------------------------------------------------------------------
ETL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ETL_DIR not in sys.path:
    sys.path.append(ETL_DIR)

from db_connection.db_connect import get_connection  # noqa: E402

# -----------------------------------------------------------------------------
# Constants (column names)
# -----------------------------------------------------------------------------
COL_FUND = "fund_name"
COL_ASSET = "asset_class_name"
COL_SECTOR = "sector"
COL_COUNTRY = "listed_country"
COL_VALUE = "value_aud"
COL_ITEM = "investment_item_name"

# -----------------------------------------------------------------------------
# Plotly styling: Montserrat + pastel palette
# -----------------------------------------------------------------------------
PASTEL_COLORS = [
    "#A3CEF1",  # soft sky blue
    "#90A7D9",  # purple-blue
    "#F7C59F",  # pastel peach
    "#ED8975",  # light coral
    "#96D7C6",  # mint
    "#F5E960",  # soft yellow
    "#CABAC8",  # rosy lavender
]

BASE_LAYOUT = dict(
    template="plotly_dark",
    font=dict(family="Montserrat", size=14, color="white"),
    paper_bgcolor="#111111",
    plot_bgcolor="#111111",
)


def apply_base_layout(fig):
    fig.update_layout(**BASE_LAYOUT)
    return fig


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def pg_ident(name: str) -> str:
    """Quote an identifier for PostgreSQL."""
    return '"' + str(name).replace('"', '""') + '"'


def _run_query_df(sql: str) -> pd.DataFrame:
    conn = get_connection()
    if conn is None:
        raise RuntimeError("Database connection not available")
    with conn, conn.cursor() as cur:
        cur.execute(sql)
        rows = cur.fetchall()
        cols = [desc[0] for desc in cur.description]
    return pd.DataFrame(rows, columns=cols)


def valid_value_pred(col: str) -> str:
    """SQL predicate for a numeric column that is not null / NaN."""
    c = pg_ident(col)
    return f"{c} IS NOT NULL AND {c}::text <> 'NaN'"


def is_subtotal_pred(col: str) -> str:
    """
    Treat any case/spacing variant of 'Sub Total' as a subtotal row.
    E.g. 'Sub Total', 'SUB TOTAL', 'sub total'.
    """
    return f"UPPER(TRIM({pg_ident(col)})) = 'SUB TOTAL'"


def normalize_asset_class(name: str) -> str:
    """Normalize asset class names: trim + collapse case + title-case."""
    if not isinstance(name, str):
        return name
    return name.strip().title()


def distinct_values(table_name: str, col: str, fund: str) -> List[str]:
    """Distinct values of a column for a given fund."""
    safe_fund = fund.replace("'", "''")
    sql = f"""
        SELECT DISTINCT TRIM({pg_ident(col)}) AS val
        FROM {pg_ident(table_name)}
        WHERE TRIM({pg_ident(COL_FUND)}) = '{safe_fund}'
          AND TRIM({pg_ident(col)}) IS NOT NULL
          AND TRIM({pg_ident(col)}) <> ''
        ORDER BY val
    """
    df = _run_query_df(sql)
    return df["val"].dropna().astype(str).tolist()


def aggregate_sum(
    table_name: str,
    dim_col: str,
    fund: str,
    extra_filters: Optional[Dict[str, List[str]]] = None,
    topn: int = 10,
) -> pd.DataFrame:
    """
    Sum of value_aud grouped by a dimension for a given fund.
    Excludes subtotal rows (any case variant of 'Sub Total').
    """
    safe_fund = fund.replace("'", "''")

    where_clauses = [
        f"TRIM({pg_ident(COL_FUND)}) = '{safe_fund}'",
        valid_value_pred(COL_VALUE),
        f"NOT {is_subtotal_pred(COL_ITEM)}",
    ]

    if extra_filters:
        for col, vals in extra_filters.items():
            if vals:
                escaped = [v.replace("'", "''") for v in vals]
                in_list = ",".join(f"'{v}'" for v in escaped)
                where_clauses.append(f"TRIM({pg_ident(col)}) IN ({in_list})")

    where_sql = " AND ".join(where_clauses)

    sql = f"""
        SELECT
            TRIM({pg_ident(dim_col)}) AS label,
            SUM({pg_ident(COL_VALUE)}) AS total
        FROM {pg_ident(table_name)}
        WHERE {where_sql}
        GROUP BY TRIM({pg_ident(dim_col)})
        ORDER BY total DESC
        LIMIT {int(topn)}
    """
    df = _run_query_df(sql)
    if not df.empty:
        df["total"] = pd.to_numeric(df["total"], errors="coerce").fillna(0.0)
    return df


def fund_totals_all_funds(table_name: str) -> pd.DataFrame:
    """Total value per fund across all rows except subtotal rows (any case)."""
    sql = f"""
        SELECT
            TRIM({pg_ident(COL_FUND)}) AS fund,
            COALESCE(
                SUM({pg_ident(COL_VALUE)})
                FILTER (
                    WHERE {valid_value_pred(COL_VALUE)}
                      AND NOT {is_subtotal_pred(COL_ITEM)}
                ),
                0
            ) AS total,
            COUNT(*) AS rows_count,
            SUM(CASE
                    WHEN {valid_value_pred(COL_VALUE)}
                     AND NOT {is_subtotal_pred(COL_ITEM)}
                    THEN 1 ELSE 0
                END) AS valid_values
        FROM {pg_ident(table_name)}
        GROUP BY TRIM({pg_ident(COL_FUND)})
        ORDER BY total DESC
    """
    df = _run_query_df(sql)
    if not df.empty:
        df["total"] = pd.to_numeric(df["total"], errors="coerce").fillna(0.0)
    return df


def asset_class_totals_all_funds(table_name: str) -> pd.DataFrame:
    """
    Totals per (fund, asset_class) using only subtotal rows.
    Any case of 'Sub Total' is treated as subtotal.
    """
    sql = f"""
        SELECT
            TRIM({pg_ident(COL_FUND)})  AS fund,
            TRIM({pg_ident(COL_ASSET)}) AS asset_class,
            SUM({pg_ident(COL_VALUE)}) FILTER (
                WHERE {is_subtotal_pred(COL_ITEM)}
                  AND {valid_value_pred(COL_VALUE)}
            ) AS total
        FROM {pg_ident(table_name)}
        GROUP BY TRIM({pg_ident(COL_FUND)}), TRIM({pg_ident(COL_ASSET)})
        HAVING SUM({pg_ident(COL_VALUE)}) FILTER (
                   WHERE {is_subtotal_pred(COL_ITEM)}
                     AND {valid_value_pred(COL_VALUE)}
               ) IS NOT NULL
        ORDER BY fund, asset_class
    """
    df = _run_query_df(sql)
    if not df.empty:
        df["total"] = pd.to_numeric(df["total"], errors="coerce").fillna(0.0)
    return df


# =============================================================================
# Main render() function
# =============================================================================
def render():
    """Render Linh's dashboard inside the 'Linh' tab."""
    st.header("Linh — Superfund dashboard")
    refresh_placeholder = st.empty()

    # ---------- SIDEBAR ----------
    with st.sidebar:
        st.markdown("---")
        st.header("All funds comparison")

        # Table name used by *all* queries
        table = st.text_input("Table name", value="final_data", help="Main table/view name")

        # Load fund list once (used for both all-funds + per-fund sections)
        try:
            sql_funds = f"""
                SELECT DISTINCT TRIM({pg_ident(COL_FUND)}) AS fund_name
                FROM {pg_ident(table)}
                ORDER BY fund_name
            """
            df_funds = _run_query_df(sql_funds)
            fund_opts = df_funds["fund_name"].dropna().astype(str).tolist()
        except Exception as e:
            st.error(f"Could not load funds from '{table}'. Error: {e}")
            return {"placeholder": refresh_placeholder}

        compare_funds = st.multiselect(
            "Funds to include (all-funds charts)",
            fund_opts,
            default=fund_opts,
        )

        st.markdown("---")
        st.header("Pick a fund")

        chosen_fund = st.selectbox(
            "Fund (for single-fund charts)",
            ["(Pick a fund)"] + fund_opts,
            index=0,
        )

        row_limit = st.number_input(
            "Preview row limit",
            min_value=10,
            max_value=100_000,
            value=2_000,
            step=100,
        )

    # ---------- Cross-fund totals (leaf rows only, excluding all subtotals) ----------
    st.subheader("Total Value (AUD) by Fund (leaf rows only)")
    try:
        g_all = fund_totals_all_funds(table)
    except Exception as e:
        st.error(f"Error loading cross-fund totals: {e}")
        return {"placeholder": refresh_placeholder}

    # Apply sidebar "Funds to include" filter
    if compare_funds:
        g_all = g_all[g_all["fund"].isin(compare_funds)]

    if g_all.empty:
        st.info("No data available for cross-fund comparison with the current selection.")
    else:
        st.caption("Debug totals: rows & valid numeric values per fund (leaf rows only)")
        st.dataframe(g_all, use_container_width=True)
        fig_total = px.bar(
            g_all,
            x="fund",
            y="total",
            color_discrete_sequence=PASTEL_COLORS,
        )
        fig_total = apply_base_layout(fig_total)
        st.plotly_chart(fig_total, use_container_width=True)

    # ---------- All-funds asset class comparison (100% STACKED, ALL ASSET CLASSES) ----------
    st.subheader("Asset Class Totals by Fund — via 'Sub Total' rows")
    try:
        df_ac = asset_class_totals_all_funds(table)
    except Exception as e:
        st.error(f"Error loading asset-class comparison: {e}")
        df_ac = pd.DataFrame()

    if df_ac.empty:
        st.info("No 'Sub Total' rows found for asset-class comparison.")
    else:
        # Filter to selected funds for all-funds chart
        if compare_funds:
            df_ac = df_ac[df_ac["fund"].isin(compare_funds)]

        if df_ac.empty:
            st.info("No data for selected funds.")
        else:
            # Normalize asset class names to prevent duplicates
            df_plot = df_ac.copy()
            df_plot["asset_class"] = df_plot["asset_class"].apply(normalize_asset_class)

            # 100% stacked: compute % share per fund
            df_plot["share"] = (
                df_plot["total"]
                / df_plot.groupby("fund")["total"].transform("sum")
                * 100.0
            )

            # Determine order of funds
            if compare_funds:
                fund_order = compare_funds
            else:
                fund_order = (
                    df_plot.groupby("fund")["total"]
                    .sum()
                    .sort_values(ascending=False)
                    .index
                    .tolist()
                )

            # Order asset classes by their average share (largest at bottom)
            asset_order = (
                df_plot.groupby("asset_class")["share"]
                .mean()
                .sort_values(ascending=False)
                .index
                .tolist()
            )

            fig_cmp = px.bar(
                df_plot,
                x="fund",
                y="share",
                color="asset_class",
                barmode="stack",
                category_orders={
                    "fund": fund_order,
                    "asset_class": asset_order,
                },
                hover_data={
                    "total": ":,.0f",   # AUD
                    "share": ":.2f",    # %
                },
                labels={
                    "share": "Allocation (%)",
                    "total": "Value (AUD)",
                    "fund": "Fund",
                    "asset_class": "Asset Class",
                },
                color_discrete_sequence=PASTEL_COLORS,
            )
            # Apply base layout first
            fig_cmp = apply_base_layout(fig_cmp)
            # Then move legend to the right side
            fig_cmp.update_layout(
                xaxis_title="Fund",
                yaxis_title="Allocation (% of fund)",
                legend_title_text="Asset Class",
                legend=dict(
                    orientation="v",
                    yanchor="top",
                    y=1,
                    xanchor="left",
                    x=1.02,   # a bit to the right of the plotting area
                ),
                margin=dict(r=220),  # extra right margin for the legend
            )
            st.caption(
                "Each bar = 100% of the fund. Hover a segment to see its exact AUD value and % allocation."
            )
            st.plotly_chart(fig_cmp, use_container_width=True)

    # ---------- Stop if no fund chosen for single-fund section ----------
    if chosen_fund == "(Pick a fund)":
        st.warning("Pick a fund in the sidebar to see preview and per-fund charts below.")
        return {"placeholder": refresh_placeholder}

    # ---------- Preview table for chosen fund ----------
    st.subheader(f"Preview — {chosen_fund}")
    safe_fund = chosen_fund.replace("'", "''")
    sql_prev = f"""
        SELECT *
        FROM {pg_ident(table)}
        WHERE TRIM({pg_ident(COL_FUND)}) = '{safe_fund}'
        ORDER BY 1
        LIMIT {int(row_limit)}
    """
    try:
        df_prev = _run_query_df(sql_prev)
    except Exception as e:
        st.error(f"Error loading preview for {chosen_fund}: {e}")
        df_prev = pd.DataFrame()

    st.dataframe(df_prev, use_container_width=True)

    # ---------- Per-fund charts with smart toggles (no sidebar filters) ----------

    # 1) Asset class pie with toggle
    g = aggregate_sum(
        table,
        COL_ASSET,
        chosen_fund,
        None,
        topn=50,
    )
    st.subheader(f"Asset Class Allocation — {chosen_fund}")

    if g.empty:
        st.info("No data for this fund.")
    else:
        total_sum = g["total"].sum()
        g["percent"] = g["total"] / total_sum * 100

        display_mode_asset = st.radio(
            "Asset class display mode:",
            ["Allocation (%)", "Value (AUD)"],
            horizontal=True,
            key="asset_mode",
        )

        if display_mode_asset == "Allocation (%)":
            top5 = g.nlargest(5, "total")["label"].tolist()
            text_labels = [
                f"{lbl} {pct:.1f}%"
                if lbl in top5 else ""
                for lbl, pct in zip(g["label"], g["percent"])
            ]
        else:
            top5 = g.nlargest(5, "total")["label"].tolist()
            text_labels = [
                f"{lbl} {val:,.0f}"
                if lbl in top5 else ""
                for lbl, val in zip(g["label"], g["total"])
            ]

        fig_asset = px.pie(
            g,
            names="label",
            values="total",
            color_discrete_sequence=PASTEL_COLORS,
        )
        fig_asset.update_traces(
            customdata=g["percent"],
            hovertemplate=(
                "Asset Class=%{label}<br>"
                "Value (AUD)=%{value:,.0f}<br>"
                "Allocation (%)=%{customdata:.2f}%"
                "<extra></extra>"
            ),
            text=text_labels,
            textinfo="text",
            textposition="outside",
            automargin=True,
        )
        fig_asset = apply_base_layout(fig_asset)
        st.caption("Hover a slice to see its % of the fund and AUD value.")
        st.plotly_chart(fig_asset, use_container_width=True)

    # 2) Sector bar with toggle
    g = aggregate_sum(
        table,
        COL_SECTOR,
        chosen_fund,
        None,
        topn=50,
    )
    st.subheader(f"Sector — {chosen_fund}")

    if g.empty:
        st.info("No data for this fund.")
    else:
        total_sum = g["total"].sum()
        g["percent"] = g["total"] / total_sum * 100

        display_mode_sector = st.radio(
            "Sector display mode:",
            ["Allocation (%)", "Value (AUD)"],
            horizontal=True,
            key="sector_mode",
        )

        if display_mode_sector == "Allocation (%)":
            y_col = "percent"
            y_title = "Allocation (%)"
        else:
            y_col = "total"
            y_title = "Value (AUD)"

        fig_sec = px.bar(
            g,
            x="label",
            y=y_col,
            hover_data={
                "percent": ":.2f",
                "total": ":,.0f",
            },
            labels={
                "label": "Sector",
                "percent": "Allocation (%)",
                "total": "Value (AUD)",
            },
            color_discrete_sequence=PASTEL_COLORS,
        )
        max_y = g[y_col].max() if not g.empty else 0
        fig_sec.update_layout(
            yaxis_title=y_title,
            yaxis=dict(range=[0, max_y * 1.1] if max_y > 0 else None),
        )
        fig_sec = apply_base_layout(fig_sec)
        st.caption("Hover bars to see both % of fund and AUD value.")
        st.plotly_chart(fig_sec, use_container_width=True)

    # 3) Country bar with toggle (Top 10 countries by value)
    g_full = aggregate_sum(
        table,
        COL_COUNTRY,
        chosen_fund,
        None,
        topn=10_000,
    )
    st.subheader(f"Listed Country — {chosen_fund}")

    if g_full.empty:
        st.info("No data for this fund.")
    else:
        total_sum_full = g_full["total"].sum()
        g_full["percent"] = g_full["total"] / total_sum_full * 100

        g = g_full.sort_values("total", ascending=False).head(10)

        display_mode_country = st.radio(
            "Country display mode:",
            ["Allocation (%)", "Value (AUD)"],
            horizontal=True,
            key="country_mode",
        )

        if display_mode_country == "Allocation (%)":
            y_col = "percent"
            y_title = "Allocation (%)"
        else:
            y_col = "total"
            y_title = "Value (AUD)"

        fig_cty = px.bar(
            g,
            x="label",
            y=y_col,
            hover_data={
                "percent": ":.2f",
                "total": ":,.0f",
            },
            labels={
                "label": "Country",
                "percent": "Allocation (%)",
                "total": "Value (AUD)",
            },
            color_discrete_sequence=PASTEL_COLORS,
        )
        max_y_cty = g[y_col].max() if not g.empty else 0
        fig_cty.update_layout(
            yaxis_title=y_title,
            yaxis=dict(range=[0, max_y_cty * 1.1] if max_y_cty > 0 else None),
        )
        fig_cty = apply_base_layout(fig_cty)
        st.caption(
            "Showing top 10 countries by value. Hover bars to see both % of fund and AUD value."
        )
        st.plotly_chart(fig_cty, use_container_width=True)

        # 3b) Australia vs International pie using the same country data
        st.subheader(f"Australia vs International — {chosen_fund}")

        aus_mask = g_full["label"].str.strip().str.upper().eq("AUSTRALIA")
        aus_total = g_full.loc[aus_mask, "total"].sum()
        rest_total = g_full.loc[~aus_mask, "total"].sum()

        if aus_total + rest_total == 0:
            st.info("No country data available to split between Australia and international.")
        else:
            df_aus = pd.DataFrame(
                {
                    "group": ["Australia", "International"],
                    "total": [aus_total, rest_total],
                }
            )
            total_sum_aus = df_aus["total"].sum()
            df_aus["percent"] = df_aus["total"] / total_sum_aus * 100

            fig_aus = px.pie(
                df_aus,
                names="group",
                values="total",
                color_discrete_sequence=PASTEL_COLORS,
            )
            fig_aus.update_traces(
                customdata=df_aus["percent"],
                hovertemplate=(
                    "Group=%{label}<br>"
                    "Value (AUD)=%{value:,.0f}<br>"
                    "Allocation (%)=%{customdata:.2f}%"
                    "<extra></extra>"
                ),
                text=[
                    f"{name} {pct:.1f}%"
                    for name, pct in zip(df_aus["group"], df_aus["percent"])
                ],
                textinfo="text",
                textposition="outside",
                automargin=True,
            )
            fig_aus = apply_base_layout(fig_aus)
            st.caption("Split of this fund between Australia and all other countries.")
            st.plotly_chart(fig_aus, use_container_width=True)

    # 4) Top holdings table within a selected asset class (for this fund)
    st.subheader(f"Top 10 holdings by value within an asset class — {chosen_fund}")

    g_asset = aggregate_sum(
        table,
        COL_ASSET,
        chosen_fund,
        None,
        topn=200,
    )

    if g_asset.empty:
        st.info("No asset classes found for this fund.")
    else:
        asset_classes = g_asset["label"].astype(str).tolist()

        selected_asset_class = st.selectbox(
            "Asset class for detailed holdings",
            asset_classes,
            key="asset_class_top_holdings",
        )

        if selected_asset_class:
            extra_filters = {COL_ASSET: [selected_asset_class]}

            g_items = aggregate_sum(
                table,
                COL_ITEM,
                chosen_fund,
                extra_filters,
                topn=10,
            )

            if g_items.empty:
                st.info("No holdings found for this asset class in this fund.")
            else:
                df_top_display = g_items.copy()
                df_top_display.rename(
                    columns={
                        "label": "Investment Item Name",
                        "total": "Value (AUD)",
                    },
                    inplace=True,
                )
                df_top_display = df_top_display.sort_values(
                    "Value (AUD)", ascending=False
                )

                st.caption(
                    f"Showing top 10 investment items in '{selected_asset_class}' for {chosen_fund}, ordered by Value (AUD)."
                )
                st.dataframe(df_top_display, use_container_width=True)

    return {"placeholder": refresh_placeholder}

