from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


APP_DIR = Path(__file__).resolve().parent


def find_data_file() -> Path:
    preferred = APP_DIR / "Amazing_Mart_Dataset_1.xlsx"
    if preferred.exists():
        return preferred

    matches = sorted(APP_DIR.glob("Amazing_Mart_Dataset_1*.xlsx"))
    if len(matches) == 1:
        return matches[0]

    if not matches:
        raise FileNotFoundError(
            "The Excel dataset was not found. Keep Amazing_Mart_Dataset_1.xlsx "
            "in the same folder as this Python file."
        )

    raise FileNotFoundError(
        "Multiple Amazing Mart Excel files were found. Keep only one dataset in the app folder."
    )


DATA_FILE = find_data_file()

NAVY = "#183153"
BLUE = "#2166F3"
SIDEBAR_BLUE = "#173B74"
TEAL = "#0C8F8F"
GREEN = "#2AA873"
GOLD = "#F2B134"
CORAL = "#E85D5D"
INK = "#102542"
MUTED = "#5F7087"
GRID = "#D7DEE8"
PAPER = "#FFFFFF"
BG = "#F3F7FB"

DISCOUNT_LABELS = ["0%", "1–10%", "11–20%", "21–30%", "31–50%", "51%+"]
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


st.set_page_config(
    page_title="Amazing Mart Dashboard",
    page_icon="AM",
    layout="wide",
    initial_sidebar_state="expanded",
)


st.markdown(
    """
    <style>
      @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600;700&family=Playfair+Display:wght@700&display=swap');

      :root {
        --navy: #183153;
        --blue: #2166F3;
        --teal: #0C8F8F;
        --ink: #102542;
        --muted: #5F7087;
        --line: #D7DEE8;
        --bg: #F3F7FB;
        --paper: #FFFFFF;
      }
      html, body, [class*="css"] {
        font-family: "DM Sans", sans-serif;
      }
      .stApp {
        background:
          radial-gradient(circle at 8% 0%, rgba(142,197,255,.24), transparent 24rem),
          radial-gradient(circle at 98% 100%, rgba(12,143,143,.12), transparent 24rem),
          var(--bg);
      }
      [data-testid="stSidebar"] {
        background: #173B74;
      }
      [data-testid="stSidebar"] * {
        color: #F8FBFF;
      }
      [data-testid="stSidebar"] label {
        font-weight: 600;
      }
      [data-testid="stSidebar"] [data-baseweb="select"] > div {
        background: rgba(255,255,255,.10);
        border-color: rgba(255,255,255,.18);
      }
      [data-testid="stSidebar"] .stDownloadButton button {
        width: 100%;
        border: 1px solid rgba(255,255,255,.28);
        background: rgba(255,255,255,.10);
        color: white;
      }
      .block-container {
        max-width: 1500px;
        padding-top: 1.4rem;
        padding-bottom: 2rem;
      }
      h1, h2, h3 {
        color: var(--navy);
      }
      h1 {
        font-family: "Playfair Display", serif;
        letter-spacing: -.02em;
      }
      div[data-testid="stMetric"] {
        background: linear-gradient(180deg, rgba(230,238,252,.95), rgba(255,255,255,.98));
        border: 1px solid rgba(16,37,66,.08);
        border-radius: 18px;
        padding: 1rem 1.1rem;
        box-shadow: 0 10px 28px rgba(24,49,83,.07);
      }
      div[data-testid="stMetric"] * {
        color: #102542 !important;
      }
      div[data-testid="stMetricLabel"] {
        color: #102542 !important;
        font-size: .76rem;
        font-weight: 700;
        letter-spacing: .08em;
        text-transform: uppercase;
      }
      div[data-testid="stMetricLabel"] * {
        color: #102542 !important;
      }
      div[data-testid="stMetricValue"] {
        color: var(--navy) !important;
        font-weight: 700;
      }
      div[data-testid="stPlotlyChart"] {
        background: var(--paper);
        border: 1px solid rgba(16,37,66,.08);
        border-radius: 20px;
        box-shadow: 0 10px 28px rgba(24,49,83,.07);
        padding: .35rem;
      }
      .view-kicker {
        color: var(--teal);
        font-size: .72rem;
        font-weight: 700;
        letter-spacing: .14em;
        text-transform: uppercase;
        margin-bottom: .15rem;
      }
      .view-title {
        color: var(--navy);
        font-family: "Playfair Display", serif;
        font-size: 2.15rem;
        line-height: 1.1;
        margin: 0 0 1rem 0;
      }
      .sidebar-brand {
        color: white;
        font-family: "Playfair Display", serif;
        font-size: 1.65rem;
        line-height: 1.05;
        margin: .25rem 0 1.1rem;
      }
      .sidebar-kicker {
        color: #9FC4FF;
        font-size: .68rem;
        font-weight: 700;
        letter-spacing: .14em;
        text-transform: uppercase;
      }
      .chart-spacer {
        height: .2rem;
      }
      button[kind="secondary"] {
        border-radius: 999px;
      }
      div[data-testid="stSegmentedControl"] button {
        border-radius: 999px;
        font-weight: 700;
      }
      @media (max-width: 800px) {
        .block-container { padding-left: .8rem; padding-right: .8rem; }
        .view-title { font-size: 1.7rem; }
      }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def load_data(path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    orders = pd.read_excel(path, sheet_name="ListOfOrders")
    lines = pd.read_excel(path, sheet_name="OrderBreakdown")
    targets = pd.read_excel(path, sheet_name="SalesTargets")

    data = lines.merge(orders, on="Order ID", how="left", validate="many_to_one")
    data["Order Date"] = pd.to_datetime(data["Order Date"])
    data["Year"] = data["Order Date"].dt.year.astype(int)
    data["Month"] = data["Order Date"].dt.month.astype(int)
    data["Month Start"] = data["Order Date"].dt.to_period("M").dt.to_timestamp()
    data["Loss"] = data["Profit"] < 0
    data["Discount Band"] = pd.cut(
        data["Discount"],
        bins=[-0.001, 0, 0.10, 0.20, 0.30, 0.50, np.inf],
        labels=DISCOUNT_LABELS,
        include_lowest=True,
    )

    targets["Month of Order Date"] = pd.to_datetime(targets["Month of Order Date"])
    targets["Year"] = targets["Month of Order Date"].dt.year.astype(int)
    targets["Month"] = targets["Month of Order Date"].dt.month.astype(int)
    return data, targets


def euro(value: float, digits: int = 0) -> str:
    sign = "-" if value < 0 else ""
    value = abs(value)
    if value >= 1_000_000:
        return f"{sign}€{value / 1_000_000:.2f}M"
    if value >= 1_000:
        return f"{sign}€{value / 1_000:.{digits}f}K"
    return f"{sign}€{value:,.0f}"


def pct(value: float, digits: int = 1) -> str:
    return f"{value * 100:.{digits}f}%"


def style_figure(fig: go.Figure, height: int = 380) -> go.Figure:
    fig.update_layout(
        height=height,
        margin=dict(l=24, r=24, t=58, b=28),
        paper_bgcolor=PAPER,
        plot_bgcolor=PAPER,
        font=dict(family="DM Sans, sans-serif", color=INK, size=12),
        title=dict(font=dict(family="DM Sans, sans-serif", color=NAVY, size=19), x=0.01, xanchor="left"),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            bgcolor="rgba(255,255,255,0)",
            font=dict(color=NAVY, size=12),
            title=dict(font=dict(color=NAVY, size=12)),
        ),
        hoverlabel=dict(bgcolor=NAVY, font_color="white", font_family="DM Sans"),
    )
    fig.update_xaxes(showgrid=False, linecolor=GRID, tickfont_color=MUTED, title_font_color=MUTED)
    fig.update_yaxes(gridcolor="#EAF0F6", zerolinecolor=GRID, tickfont_color=MUTED, title_font_color=MUTED)
    return fig


def chart(fig: go.Figure, key: str) -> None:
    st.plotly_chart(
        fig,
        width="stretch",
        config={
            "displaylogo": False,
            "modeBarButtonsToRemove": ["lasso2d", "select2d"],
            "toImageButtonOptions": {"format": "png", "scale": 2},
        },
        key=key,
    )


def view_heading(kicker: str, title: str) -> None:
    st.markdown(
        f'<div class="view-kicker">{kicker}</div>'
        f'<div class="view-title">{title}</div>',
        unsafe_allow_html=True,
    )


def kpis(frame: pd.DataFrame) -> None:
    sales = frame["Sales"].sum()
    profit = frame["Profit"].sum()
    margin = profit / sales if sales else 0
    loss_rate = frame["Loss"].mean() if len(frame) else 0
    columns = st.columns(4)
    columns[0].metric("Sales", euro(sales))
    columns[1].metric("Profit", euro(profit))
    columns[2].metric("Profit Margin", pct(margin))
    columns[3].metric("Unprofitable Lines", pct(loss_rate))


def filtered_frame(data: pd.DataFrame) -> pd.DataFrame:
    frame = data.copy()
    if selected_regions:
        frame = frame[frame["Region"].isin(selected_regions)]
    if selected_segments:
        frame = frame[frame["Segment"].isin(selected_segments)]
    if selected_categories:
        frame = frame[frame["Category"].isin(selected_categories)]
    if selected_years:
        frame = frame[frame["Year"].isin(selected_years)]
    return frame


def render_overview(frame: pd.DataFrame) -> None:
    view_heading("Executive View", "Amazing Mart Performance")
    kpis(frame)
    st.write("")

    monthly = (
        frame.groupby("Month Start", as_index=False)
        .agg(Sales=("Sales", "sum"), Profit=("Profit", "sum"))
        .sort_values("Month Start")
    )
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=monthly["Month Start"], y=monthly["Sales"], name="Sales",
        mode="lines", line=dict(color=BLUE, width=3),
        hovertemplate="%{x|%b %Y}<br>Sales: €%{y:,.0f}<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=monthly["Month Start"], y=monthly["Profit"], name="Profit",
        mode="lines", line=dict(color=TEAL, width=3), yaxis="y2",
        hovertemplate="%{x|%b %Y}<br>Profit: €%{y:,.0f}<extra></extra>",
    ))
    fig.update_layout(
        title="Sales and Profit Trend",
        yaxis=dict(title="Sales (€)", tickprefix="€", separatethousands=True),
        yaxis2=dict(title="Profit (€)", overlaying="y", side="right", showgrid=False, tickprefix="€", separatethousands=True),
        hovermode="x unified",
    )
    chart(style_figure(fig, 420), "overview_trend")

    left, right = st.columns(2)
    with left:
        category = frame.groupby("Category", as_index=False).agg(
            Sales=("Sales", "sum"), Profit=("Profit", "sum")
        )
        category["Margin"] = category["Profit"] / category["Sales"]
        category = category.sort_values("Margin")
        fig = px.bar(
            category, x="Category", y="Margin", color="Margin",
            color_continuous_scale=[CORAL, GOLD, GREEN],
            title="Profit Margin by Category",
            text=category["Margin"].map(lambda x: pct(x)),
        )
        fig.update_traces(textposition="outside", hovertemplate="%{x}<br>Margin: %{y:.1%}<extra></extra>")
        fig.update_layout(coloraxis_showscale=False)
        fig.update_yaxes(tickformat=".0%")
        chart(style_figure(fig, 360), "overview_category")
    with right:
        region = frame.groupby("Region", as_index=False).agg(
            Sales=("Sales", "sum"), Profit=("Profit", "sum")
        )
        region["Margin"] = region["Profit"] / region["Sales"]
        region = region.sort_values("Margin")
        fig = px.bar(
            region, x="Region", y="Margin", color="Margin",
            color_continuous_scale=[CORAL, GOLD, GREEN],
            title="Profit Margin by Region",
            text=region["Margin"].map(lambda x: pct(x)),
        )
        fig.update_traces(textposition="outside", hovertemplate="%{x}<br>Margin: %{y:.1%}<extra></extra>")
        fig.update_layout(coloraxis_showscale=False)
        fig.update_yaxes(tickformat=".0%")
        chart(style_figure(fig, 360), "overview_region")


def render_discount(frame: pd.DataFrame) -> None:
    view_heading("Question 1", "Discount Risk")
    kpis(frame)
    st.write("")

    bands = (
        frame.groupby("Discount Band", observed=False)
        .agg(Loss_Rate=("Loss", "mean"), Average_Profit=("Profit", "mean"), Lines=("Profit", "size"))
        .reindex(DISCOUNT_LABELS)
        .reset_index()
    )
    left, right = st.columns(2)
    with left:
        colors = [GREEN if x < .2 else GOLD if x < .5 else CORAL for x in bands["Loss_Rate"].fillna(0)]
        fig = go.Figure(go.Bar(
            x=bands["Discount Band"], y=bands["Loss_Rate"], marker_color=colors,
            text=bands["Loss_Rate"].map(lambda x: pct(x, 0) if pd.notna(x) else ""),
            textposition="outside",
            hovertemplate="%{x}<br>Loss rate: %{y:.1%}<extra></extra>",
        ))
        fig.update_layout(title="Probability of Loss by Discount Band")
        fig.update_yaxes(tickformat=".0%", range=[0, max(1.05, bands["Loss_Rate"].max() * 1.18)])
        chart(style_figure(fig, 370), "discount_loss")
    with right:
        colors = [CORAL if x < 0 else GREEN for x in bands["Average_Profit"].fillna(0)]
        fig = go.Figure(go.Bar(
            x=bands["Discount Band"], y=bands["Average_Profit"], marker_color=colors,
            text=bands["Average_Profit"].map(lambda x: euro(x) if pd.notna(x) else ""),
            textposition="outside",
            hovertemplate="%{x}<br>Average profit: €%{y:,.0f}<extra></extra>",
        ))
        fig.update_layout(title="Average Profit by Discount Band")
        fig.add_hline(y=0, line_dash="dot", line_color=MUTED)
        chart(style_figure(fig, 370), "discount_profit")

    display = frame if len(frame) <= 4500 else frame.sample(4500, random_state=42).copy()
    display["Outcome"] = np.where(display["Loss"], "Loss", "Profit")
    fig = px.scatter(
        display,
        x="Discount",
        y="Profit",
        color="Outcome",
        color_discrete_map={"Profit": GREEN, "Loss": CORAL},
        hover_data=["Category", "Sub-Category", "Region", "Sales"],
        opacity=.56,
        title="Discount vs. Profit",
        labels={"Outcome": "Profit and Loss"},
    )
    fig.update_traces(marker=dict(size=6))
    fig.update_layout(legend_title_text="Profit and Loss")
    fig.update_xaxes(tickformat=".0%")
    fig.add_hline(y=0, line_dash="dot", line_color=MUTED)
    chart(style_figure(fig, 480), "discount_scatter")

    x = np.linspace(0, .60, 121)
    probability = 1 / (1 + np.exp(-(-3.61 + 14.93 * x)))
    fig = go.Figure(go.Scatter(
        x=x, y=probability, mode="lines", line=dict(color=NAVY, width=4),
        hovertemplate="Discount: %{x:.0%}<br>Predicted loss: %{y:.1%}<extra></extra>",
    ))
    fig.add_hline(y=.5, line_dash="dash", line_color=MUTED)
    fig.add_vline(x=.242, line_dash="dash", line_color=CORAL)
    fig.update_layout(title="Predicted Probability of Loss")
    fig.update_xaxes(tickformat=".0%", title="Discount")
    fig.update_yaxes(tickformat=".0%", title="Probability", range=[0, 1.05])
    chart(style_figure(fig, 350), "discount_model")


def render_products(frame: pd.DataFrame) -> None:
    view_heading("Question 2", "Product Performance")
    kpis(frame)
    st.write("")

    metric = st.segmented_control(
        "Metric",
        options=["Profit", "Profit Margin", "Sales"],
        default="Profit",
        key="product_metric",
    )
    metric = metric or "Profit"

    category = frame.groupby("Category", as_index=False).agg(
        Sales=("Sales", "sum"), Profit=("Profit", "sum")
    )
    category["Profit Margin"] = category["Profit"] / category["Sales"]
    subcategory = frame.groupby("Sub-Category", as_index=False).agg(
        Sales=("Sales", "sum"), Profit=("Profit", "sum")
    )
    subcategory["Profit Margin"] = subcategory["Profit"] / subcategory["Sales"]

    left, right = st.columns([.85, 1.15])
    with left:
        category = category.sort_values(metric)
        colors = [CORAL if x < 0 else GREEN for x in category[metric]]
        fig = go.Figure(go.Bar(
            x=category["Category"], y=category[metric], marker_color=colors,
            text=[pct(v) if metric == "Profit Margin" else euro(v) for v in category[metric]],
            textposition="outside",
            hovertemplate=f"%{{x}}<br>{metric}: %{{y:,.2f}}<extra></extra>",
        ))
        fig.update_layout(title=f"{metric} by Category")
        if metric == "Profit Margin":
            fig.update_yaxes(tickformat=".0%")
        fig.add_hline(y=0, line_dash="dot", line_color=MUTED)
        chart(style_figure(fig, 430), "product_category")
    with right:
        subcategory = subcategory.sort_values(metric)
        colors = [CORAL if x < 0 else TEAL for x in subcategory[metric]]
        fig = go.Figure(go.Bar(
            y=subcategory["Sub-Category"], x=subcategory[metric], orientation="h",
            marker_color=colors,
            hovertemplate=f"%{{y}}<br>{metric}: %{{x:,.2f}}<extra></extra>",
        ))
        fig.update_layout(title=f"{metric} by Sub-Category")
        if metric == "Profit Margin":
            fig.update_xaxes(tickformat=".0%")
        fig.add_vline(x=0, line_dash="dot", line_color=MUTED)
        chart(style_figure(fig, 520), "product_subcategory")

    heat = frame.pivot_table(
        index="Category", columns="Region", values="Profit", aggfunc="sum", fill_value=0
    )
    fig = px.imshow(
        heat,
        text_auto=",.0f",
        color_continuous_scale=["#F4D7D7", "#F7F2E8", "#4FAF7C"],
        color_continuous_midpoint=0,
        aspect="auto",
        title="Category-Region Profit Heatmap",
        labels=dict(color="Profit (€)"),
    )
    fig.update_traces(hovertemplate="%{y} · %{x}<br>Profit: €%{z:,.0f}<extra></extra>")
    fig.update_layout(coloraxis_colorbar=dict(title=dict(text="Profit (€)", font=dict(color=NAVY)), tickfont=dict(color=NAVY)))
    chart(style_figure(fig, 350), "product_heatmap")


def render_geography(frame: pd.DataFrame) -> None:
    view_heading("Question 3", "Geographic Performance")
    kpis(frame)
    st.write("")

    metric = st.segmented_control(
        "Metric",
        options=["Profit Margin", "Profit", "Sales", "Loss Rate"],
        default="Profit Margin",
        key="geo_metric",
    )
    metric = metric or "Profit Margin"

    grouped = frame.groupby(["Region", "Country"], as_index=False).agg(
        Sales=("Sales", "sum"),
        Profit=("Profit", "sum"),
        Loss_Rate=("Loss", "mean"),
    )
    grouped["Profit Margin"] = grouped["Profit"] / grouped["Sales"]
    grouped = grouped.rename(columns={"Loss_Rate": "Loss Rate"})

    region = frame.groupby("Region", as_index=False).agg(
        Sales=("Sales", "sum"),
        Profit=("Profit", "sum"),
        Loss_Rate=("Loss", "mean"),
    )
    region["Profit Margin"] = region["Profit"] / region["Sales"]
    region = region.rename(columns={"Loss_Rate": "Loss Rate"}).sort_values(metric)

    map_values = grouped[metric].astype(float)
    if metric in {"Profit", "Profit Margin"}:
        map_limit = max(abs(map_values.min()), abs(map_values.max()), .001)
        color_min, color_max = -map_limit, map_limit
        color_scale = [[0, CORAL], [.5, "#FFF5DC"], [1, GREEN]]
    elif metric == "Loss Rate":
        color_min, color_max = 0, max(float(map_values.max()), .01)
        color_scale = [[0, GREEN], [.5, GOLD], [1, CORAL]]
    else:
        color_min, color_max = 0, max(float(map_values.max()), 1)
        color_scale = [[0, "#BFD7FF"], [.55, BLUE], [1, NAVY]]

    max_sales = max(float(grouped["Sales"].max()), 1)
    bubble_sizes = 12 + 34 * np.sqrt(grouped["Sales"] / max_sales)
    metric_hover = (
        grouped[metric].map(lambda value: pct(value))
        if metric in {"Profit Margin", "Loss Rate"}
        else grouped[metric].map(euro)
    )
    map_custom = np.column_stack([
        grouped["Region"],
        grouped["Sales"].map(euro),
        grouped["Profit"].map(euro),
        metric_hover,
    ])

    map_fig = go.Figure()
    map_fig.add_trace(go.Choropleth(
        locations=grouped["Country"],
        locationmode="country names",
        z=grouped[metric],
        colorscale=color_scale,
        zmin=color_min,
        zmax=color_max,
        marker_line_color="#FFFFFF",
        marker_line_width=.8,
        showscale=False,
        hoverinfo="skip",
    ))
    map_fig.add_trace(go.Scattergeo(
        locations=grouped["Country"],
        locationmode="country names",
        text=grouped["Country"],
        customdata=map_custom,
        mode="markers",
        marker=dict(
            size=bubble_sizes,
            color=grouped[metric],
            colorscale=color_scale,
            cmin=color_min,
            cmax=color_max,
            opacity=.84,
            line=dict(color="white", width=1.6),
            colorbar=dict(
                title=metric,
                thickness=12,
                len=.62,
                x=1.01,
                outlinewidth=0,
            ),
        ),
        hovertemplate=(
            "<b>%{text}</b><br>"
            "Region: %{customdata[0]}<br>"
            f"{metric}: %{{customdata[3]}}<br>"
            "Sales: %{customdata[1]}<br>"
            "Profit: %{customdata[2]}<extra></extra>"
        ),
    ))
    map_fig.update_geos(
        scope="europe",
        projection_type="mercator",
        showland=True,
        landcolor="#F8F4EA",
        showcountries=True,
        countrycolor="#B8C7D6",
        countrywidth=.7,
        showcoastlines=True,
        coastlinecolor="#9DB0C1",
        showocean=True,
        oceancolor="#E7F1F8",
        showlakes=True,
        lakecolor="#E7F1F8",
        bgcolor=PAPER,
        lataxis_range=[34, 72],
        lonaxis_range=[-13, 32],
    )
    map_fig.update_layout(
        title=f"European Performance Map · Bubble Size = Sales · Color = {metric}",
        showlegend=False,
    )
    chart(style_figure(map_fig, 590), "geo_map")

    left, right = st.columns([.7, 1.3])
    with left:
        colors = [CORAL if x < 0 else TEAL for x in region[metric]]
        if metric == "Loss Rate":
            colors = [CORAL if x > .25 else GOLD if x > .15 else GREEN for x in region[metric]]
        fig = go.Figure(go.Bar(
            x=region["Region"], y=region[metric], marker_color=colors,
            text=[pct(v) if metric in {"Profit Margin", "Loss Rate"} else euro(v) for v in region[metric]],
            textposition="outside",
            hovertemplate=f"%{{x}}<br>{metric}: %{{y:,.2f}}<extra></extra>",
        ))
        fig.update_layout(title=f"{metric} by Region")
        if metric in {"Profit Margin", "Loss Rate"}:
            fig.update_yaxes(tickformat=".0%")
        chart(style_figure(fig, 410), "geo_region")
    with right:
        country = grouped.sort_values(metric)
        colors = [CORAL if x < 0 else TEAL for x in country[metric]]
        if metric == "Loss Rate":
            colors = [CORAL if x > .25 else GOLD if x > .15 else GREEN for x in country[metric]]
        fig = go.Figure(go.Bar(
            y=country["Country"], x=country[metric], orientation="h", marker_color=colors,
            customdata=country[["Region", "Sales", "Profit"]].to_numpy(),
            hovertemplate=(
                "%{y}<br>Region: %{customdata[0]}"
                f"<br>{metric}: %{{x:,.2f}}"
                "<br>Sales: €%{customdata[1]:,.0f}"
                "<br>Profit: €%{customdata[2]:,.0f}<extra></extra>"
            ),
        ))
        fig.update_layout(title=f"{metric} by Country")
        if metric in {"Profit Margin", "Loss Rate"}:
            fig.update_xaxes(tickformat=".0%")
        fig.add_vline(x=0, line_dash="dot", line_color=MUTED)
        chart(style_figure(fig, 520), "geo_country")

    matrix = grouped.pivot(index="Country", columns="Region", values=metric)
    fig = px.imshow(
        matrix,
        text_auto=".1%" if metric in {"Profit Margin", "Loss Rate"} else ",.0f",
        color_continuous_scale=["#F4D7D7", "#F7F2E8", "#4FAF7C"],
        aspect="auto",
        title=f"Country and Region {metric}",
        labels=dict(color=metric),
    )
    fig.update_layout(coloraxis_colorbar=dict(title=dict(text=metric, font=dict(color=NAVY)), tickfont=dict(color=NAVY)))
    chart(style_figure(fig, 480), "geo_matrix")


def target_data(data: pd.DataFrame, targets: pd.DataFrame) -> pd.DataFrame:
    frame = data.copy()
    if selected_years:
        frame = frame[frame["Year"].isin(selected_years)]
    if selected_categories:
        frame = frame[frame["Category"].isin(selected_categories)]
    actual = (
        frame.groupby(["Month Start", "Category"], as_index=False)["Sales"]
        .sum()
        .rename(columns={"Month Start": "Month of Order Date", "Sales": "Actual"})
    )
    target = targets.copy()
    if selected_years:
        target = target[target["Year"].isin(selected_years)]
    if selected_categories:
        target = target[target["Category"].isin(selected_categories)]
    result = target.merge(actual, on=["Month of Order Date", "Category"], how="left")
    result["Actual"] = result["Actual"].fillna(0)
    result["Attainment"] = result["Actual"] / result["Target"]
    result["Met"] = result["Attainment"] >= 1
    return result


def render_targets(data: pd.DataFrame, targets: pd.DataFrame) -> None:
    view_heading("Question 4", "Sales Target Attainment")
    target = target_data(data, targets)

    attainment = target["Actual"].sum() / target["Target"].sum() if target["Target"].sum() else 0
    met = int(target["Met"].sum())
    periods = len(target)
    columns = st.columns(4)
    columns[0].metric("Actual Sales", euro(target["Actual"].sum()))
    columns[1].metric("Sales Target", euro(target["Target"].sum()))
    columns[2].metric("Attainment", pct(attainment))
    columns[3].metric("Targets Met", f"{met} / {periods}")
    st.write("")

    gauge = go.Figure(go.Indicator(
        mode="gauge+number",
        value=attainment,
        number=dict(valueformat=".0%", font=dict(size=52, color=NAVY)),
        gauge=dict(
            axis=dict(
                range=[0, 1.2],
                tickmode="array",
                tickvals=[0, .5, .9, 1, 1.2],
                ticktext=["0%", "50%", "90%", "100%", "120%"],
                tickcolor=MUTED,
            ),
            bar=dict(color=NAVY, thickness=.28),
            bgcolor="#F4F7FA",
            borderwidth=0,
            steps=[
                dict(range=[0, .9], color="#F6C8C8"),
                dict(range=[.9, 1], color="#F6E3A8"),
                dict(range=[1, 1.2], color="#BFE8D2"),
            ],
            threshold=dict(
                line=dict(color=CORAL, width=4),
                thickness=.8,
                value=1,
            ),
        ),
    ))
    gauge.update_layout(
        margin=dict(l=42, r=42, t=90, b=20),
        title=dict(text="Overall Target Attainment", x=0.01, xanchor="left", font=dict(size=19, color=NAVY)),
    )
    chart(style_figure(gauge, 360), "target_speedometer")

    category = target.groupby("Category", as_index=False).agg(
        Actual=("Actual", "sum"), Target=("Target", "sum"), Months_Met=("Met", "sum")
    )
    category["Attainment"] = category["Actual"] / category["Target"]

    left, right = st.columns(2)
    with left:
        colors = [GREEN if x >= 1 else GOLD if x >= .9 else CORAL for x in category["Attainment"]]
        fig = go.Figure(go.Bar(
            x=category["Category"], y=category["Attainment"], marker_color=colors,
            text=category["Attainment"].map(lambda x: pct(x, 0)), textposition="outside",
            hovertemplate="%{x}<br>Attainment: %{y:.1%}<extra></extra>",
        ))
        fig.add_hline(y=1, line_dash="dash", line_color=NAVY)
        fig.update_layout(title="Target Attainment by Category")
        fig.update_yaxes(tickformat=".0%")
        chart(style_figure(fig, 380), "target_category")
    with right:
        possible = target.groupby("Category").size().reindex(category["Category"]).to_numpy()
        fig = go.Figure(go.Bar(
            x=category["Category"], y=category["Months_Met"], marker_color=TEAL,
            text=[f"{a}/{b}" for a, b in zip(category["Months_Met"], possible)],
            textposition="outside",
            hovertemplate="%{x}<br>Periods met: %{y}<extra></extra>",
        ))
        fig.update_layout(title="Months Target Was Met")
        chart(style_figure(fig, 380), "target_met")

    target["Month Name"] = pd.Categorical(
        target["Month of Order Date"].dt.strftime("%b"),
        categories=MONTHS,
        ordered=True,
    )
    if len(selected_years) == 1:
        heat = target.pivot_table(
            index="Category", columns="Month Name", values="Attainment", observed=False
        ).reindex(columns=MONTHS)
        heat_title = f"Monthly Target Attainment · {selected_years[0]}"
    else:
        heat = target.pivot_table(
            index="Category", columns="Month Name", values="Attainment",
            aggfunc="mean", observed=False
        ).reindex(columns=MONTHS)
        heat_title = "Average Monthly Target Attainment"

    fig = px.imshow(
        heat,
        text_auto=".0%",
        color_continuous_scale=["#F4D7D7", "#F6E3A8", "#4FAF7C"],
        color_continuous_midpoint=1,
        aspect="auto",
        zmin=max(0, float(np.nanmin(heat.to_numpy())) * .95),
        zmax=max(1.1, float(np.nanmax(heat.to_numpy()))),
        title=heat_title,
        labels=dict(color="Attainment"),
    )
    fig.update_traces(hovertemplate="%{y} · %{x}<br>Attainment: %{z:.1%}<extra></extra>")
    fig.update_layout(coloraxis_colorbar=dict(title=dict(text="Attainment", font=dict(color=NAVY)), tickfont=dict(color=NAVY)))
    chart(style_figure(fig, 360), "target_heatmap")

    monthly = target.groupby("Month of Order Date", as_index=False).agg(
        Actual=("Actual", "sum"), Target=("Target", "sum")
    )
    monthly["Attainment"] = monthly["Actual"] / monthly["Target"]
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=monthly["Month of Order Date"], y=monthly["Attainment"],
        mode="lines+markers", line=dict(color=TEAL, width=3),
        marker=dict(
            color=np.where(monthly["Attainment"] >= 1, GREEN, CORAL),
            size=7,
        ),
        hovertemplate="%{x|%b %Y}<br>Attainment: %{y:.1%}<extra></extra>",
    ))
    fig.add_hline(y=1, line_dash="dash", line_color=NAVY)
    fig.update_layout(title="Monthly Attainment Trend")
    fig.update_yaxes(tickformat=".0%")
    chart(style_figure(fig, 400), "target_trend")


data, targets = load_data(DATA_FILE)

with st.sidebar:
    st.markdown('<div class="sidebar-kicker">European Superstore · 2011–2014</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-brand">Amazing Mart<br>Dashboard</div>', unsafe_allow_html=True)

    view = st.radio(
        "View",
        ["Overview", "Discount Risk", "Products", "Geography", "Targets"],
        captions=[
            "Executive summary",
            "Question 1",
            "Question 2",
            "Question 3",
            "Question 4",
        ],
    )
    st.divider()
    st.markdown("#### Filters")

    selected_years = st.multiselect(
        "Year", sorted(data["Year"].unique()), default=sorted(data["Year"].unique())
    )
    selected_regions = st.multiselect(
        "Region", sorted(data["Region"].dropna().unique()),
        default=sorted(data["Region"].dropna().unique()),
    )
    selected_segments = st.multiselect(
        "Segment", sorted(data["Segment"].dropna().unique()),
        default=sorted(data["Segment"].dropna().unique()),
    )
    selected_categories = st.multiselect(
        "Category", sorted(data["Category"].dropna().unique()),
        default=sorted(data["Category"].dropna().unique()),
    )

frame = filtered_frame(data)

with st.sidebar:
    st.caption(f"{len(frame):,} of {len(data):,} line items")
    st.download_button(
        "Download filtered data",
        data=frame.to_csv(index=False).encode("utf-8"),
        file_name="amazing_mart_filtered.csv",
        mime="text/csv",
    )

if frame.empty:
    st.warning("No data matches the selected filters.")
    st.stop()

if view == "Overview":
    render_overview(frame)
elif view == "Discount Risk":
    render_discount(frame)
elif view == "Products":
    render_products(frame)
elif view == "Geography":
    render_geography(frame)
else:
    render_targets(data, targets)
