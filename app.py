"""A small public demonstration of a private spreadsheet-analysis workflow."""

import os
from pathlib import Path

from dash import Dash, Input, Output, State, callback, dash_table, dcc, html
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from analysis import batch_summary, export_workbook, filter_records, material_pareto, metrics, monthly_summary, read_records
from demo_data import generate_workbook


ROOT = Path(__file__).resolve().parent
# Always regenerate this fictional input. No private paths, uploads or database fallback.
DATA = read_records(generate_workbook(ROOT / "demo" / "synthetic_records.xlsx"))
COLORS = ["#254761", "#50877D", "#C99B68"]
GRAPH_CONFIG = {"displayModeBar": False, "responsive": True}
MONTHS = [{"label": pd.Timestamp(f"{month}-01").strftime("%B %Y"), "value": month} for month in sorted(DATA["month"].unique())]

app = Dash(__name__, title="Costos · Portfolio demo")
server = app.server


def polish(figure, height=320):
    figure.update_layout(
        template="plotly_white", height=height, margin=dict(l=16, r=20, t=20, b=30),
        font=dict(family="Inter, Arial, sans-serif", size=12, color="#52606D"),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, title=None),
        hoverlabel=dict(bgcolor="#FFFFFF", font_size=12),
    )
    figure.update_xaxes(showgrid=False, zeroline=False)
    figure.update_yaxes(gridcolor="#E9E7E1", zeroline=False)
    return figure


def empty_chart():
    figure = go.Figure()
    figure.add_annotation(text="No records match these filters", showarrow=False, x=.5, y=.5, xref="paper", yref="paper")
    return polish(figure)


def metric_card(label, element_id, subtitle):
    return html.Div([html.P(label, className="metric-label"), html.Div(id=element_id, className="metric-value"), html.P(subtitle, className="metric-note")], className="metric-card")


app.layout = html.Div([
    html.Div([
        html.Span(className="demo-dot"),
        html.Span("Portfolio demo • synthetic data"),
        html.Span("Every record is fictional", className="demo-note"),
    ], className="demo-banner"),
    html.Main([
        html.Header([
            html.Div([
                html.Div([html.Span("C", className="brand-mark"), html.Span("COSTOS", className="brand-name")], className="brand"),
                html.P("FROM SPREADSHEETS TO A CLEARER VIEW", className="eyebrow"),
                html.H1(["A closer look at ", html.Span("the numbers.")]),
                html.P("Explore monthly and batch costs with a synthetic example from tanning and leather-processing analysis.", className="intro"),
            ]),
            html.Div([html.Span("WORKFLOW", className="eyebrow"), html.P("Excel  →  Python  →  Insights", className="workflow"), html.P("72 fictional batches · May–August 2026", className="period-note")], className="hero-note"),
        ], className="hero"),
        html.Section([
            html.Div([html.Label("CLIENT", htmlFor="client-filter"), dcc.Dropdown(id="client-filter", options=[{"label": client, "value": client} for client in sorted(DATA["client"].unique())], multi=True, placeholder="All clients")], className="filter-field"),
            html.Div([html.Label("MONTH", htmlFor="month-filter"), dcc.Dropdown(id="month-filter", options=MONTHS, multi=True, placeholder="All months")], className="filter-field"),
            html.Div([html.Button("Download Excel ↗", id="export-button", n_clicks=0), html.Span("Exports the current selection", className="export-note")], className="export-block"),
            dcc.Download(id="download"),
        ], className="filter-bar"),
        html.Div([
            metric_card("TOTAL COST", "total-cost", "Illustrative USD"),
            metric_card("COST PER KG", "cost-per-kg", "Weighted by batch size"),
            metric_card("BATCHES", "batch-count", "Distinct records in this view"),
            metric_card("BATCH WEIGHT", "batch-weight", "Each batch counted once"),
        ], className="metrics"),
        html.Section([
            html.Div([
                html.Div([html.Div([html.P("THE BIG PICTURE", className="eyebrow"), html.H2("Costs over time")]), html.Span("BY CLIENT", className="section-tag")], className="chart-heading"),
                html.P("Monthly totals from the selected records.", className="chart-caption"),
                dcc.Graph(id="monthly-chart", config=GRAPH_CONFIG, style={"height": "320px"}),
            ], className="chart-card"),
            html.Div([
                html.Div([html.Div([html.P("WHERE IT GOES", className="eyebrow"), html.H2("Spending concentration")]), html.Span("80 / 20", className="section-tag")], className="chart-heading"),
                html.P("Bars show cost; the line tracks cumulative share.", className="chart-caption"),
                dcc.Graph(id="pareto-chart", config=GRAPH_CONFIG, style={"height": "320px"}),
            ], className="chart-card"),
        ], className="charts"),
        html.Div(id="insight", className="insight"),
        html.Section([
            html.Div([html.Div([html.P("BEHIND THE CHARTS", className="eyebrow"), html.H2("Batch detail")]), html.Span(id="table-count", className="section-tag")], className="chart-heading"),
            html.P("Source rows are grouped by batch, so kilograms are never counted twice.", className="chart-caption"),
            dash_table.DataTable(
                id="batch-table", columns=[
                    {"name": "Batch", "id": "batch_id"}, {"name": "Date", "id": "date"},
                    {"name": "Client", "id": "client"}, {"name": "Weight (kg)", "id": "base_kg", "type": "numeric"},
                    {"name": "Cost (USD)", "id": "cost", "type": "numeric"},
                    {"name": "USD / kg", "id": "cost_per_kg", "type": "numeric"},
                ], page_size=6, sort_action="native", style_as_list_view=True,
                style_table={"overflowX": "auto"},
                style_cell={"fontFamily": "Inter, Arial, sans-serif", "fontSize": 13, "padding": "13px 16px", "textAlign": "left", "color": "#354B5C", "backgroundColor": "#FFFFFF", "border": "none", "borderBottom": "1px solid #EEEDE8"},
                style_header={"fontWeight": "600", "backgroundColor": "#F7F7F3", "color": "#69747D", "fontSize": 11},
                style_data_conditional=[{"if": {"row_index": "odd"}, "backgroundColor": "#FCFCFA"}],
            ),
        ], className="table-card"),
        html.Footer([html.Span("Developed primarily with Codex, using cost-analysis context to guide prompts and review outputs."), html.Span("A public demonstration of a private Costos workflow. All values are synthetic.")]),
    ], className="page"),
])


@callback(
    Output("total-cost", "children"), Output("cost-per-kg", "children"),
    Output("batch-count", "children"), Output("batch-weight", "children"),
    Output("monthly-chart", "figure"), Output("pareto-chart", "figure"),
    Output("insight", "children"), Output("batch-table", "data"), Output("table-count", "children"),
    Input("client-filter", "value"), Input("month-filter", "value"),
)
def update_dashboard(clients, months):
    selected = filter_records(DATA, clients, months)
    totals = metrics(selected)
    if selected.empty:
        return "$0", "$0.00", "0", "0 kg", empty_chart(), empty_chart(), "No records match the selected filters.", [], "0 BATCHES"
    monthly = monthly_summary(selected)
    monthly["Month"] = pd.to_datetime(monthly["month"] + "-01").dt.strftime("%b %Y")
    monthly_figure = px.bar(monthly, x="Month", y="cost", color="client", barmode="stack", color_discrete_map=dict(zip(sorted(DATA["client"].unique()), COLORS)), labels={"cost": "Cost (USD)", "client": "Client"})
    monthly_figure.update_traces(hovertemplate="%{x}<br>$%{y:,.2f}<extra>%{fullData.name}</extra>")
    monthly_figure.update_xaxes(title=None)
    monthly_figure.update_yaxes(title=None, tickprefix="$", tickformat="~s")
    polish(monthly_figure)
    pareto = material_pareto(selected)
    pareto_figure = make_subplots(specs=[[{"secondary_y": True}]])
    pareto_figure.add_trace(go.Bar(x=pareto["material"], y=pareto["cost"], marker_color=[COLORS[0] if priority else "#CCD7DC" for priority in pareto["priority"]], name="Cost", hovertemplate="%{x}<br>$%{y:,.2f}<extra></extra>"), secondary_y=False)
    pareto_figure.add_trace(go.Scatter(x=pareto["material"], y=pareto["cumulative_share"], mode="lines+markers", name="Cumulative share", line=dict(color="#B88145", width=2), marker=dict(size=6), hovertemplate="%{x}<br>%{y:.1%} of total<extra></extra>"), secondary_y=True)
    pareto_figure.add_hline(y=.8, line_dash="dot", line_color="#B88145", line_width=1, secondary_y=True)
    polish(pareto_figure)
    pareto_figure.update_yaxes(tickprefix="$", tickformat="~s", secondary_y=False)
    pareto_figure.update_yaxes(tickformat=".0%", range=[0, 1.08], showgrid=False, secondary_y=True)
    pareto_figure.update_xaxes(tickangle=-15)
    priority = pareto[pareto["priority"]]
    share = float(priority["share"].sum())
    insight = [html.Span("A useful place to start", className="insight-label"), html.Span(f"{len(priority)} of {len(pareto)} spending items account for {share:.0%} of selected costs. This describes the fictional data; it is not a savings estimate.")]
    batches = batch_summary(selected).sort_values(["date", "batch_id"], ascending=[False, True])
    batches["date"] = batches["date"].dt.strftime("%b %d, %Y")
    batches[["base_kg", "cost", "cost_per_kg"]] = batches[["base_kg", "cost", "cost_per_kg"]].round(2)
    return f"${totals['total_cost']:,.0f}", f"${totals['cost_per_kg']:,.2f}", f"{totals['batches']:,}", f"{totals['base_kg']:,.0f} kg", monthly_figure, pareto_figure, insight, batches.to_dict("records"), f"{totals['batches']} BATCHES"


@callback(Output("download", "data"), Input("export-button", "n_clicks"), State("client-filter", "value"), State("month-filter", "value"), prevent_initial_call=True)
def download_excel(_clicks, clients, months):
    return dcc.send_bytes(export_workbook(filter_records(DATA, clients, months)), "costos_synthetic_selection.xlsx")


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", "8050")), debug=False)
