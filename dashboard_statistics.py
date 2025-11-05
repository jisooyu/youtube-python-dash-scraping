import os
import pandas as pd
from dash import Dash, html, dcc, dash_table

# ============================================================
# 1️⃣ Load Excel file
# ============================================================
input_path = os.path.join(os.path.dirname(__file__), "StatisticsSummary.xlsx")

# Safely handle missing file
if not os.path.exists(input_path):
    import dash
    from dash import html
    app = dash.Dash(__name__)
    app.layout = html.Div([
        html.H3("⚠️ Data file not found."),
        html.P("Please upload or generate StatisticsSummary.xlsx in this directory.")
    ])
else:
    df = pd.read_excel(input_path, sheet_name="statistics_summary")

df = pd.read_excel(input_path, sheet_name="statistics_summary")

# ============================================================
# 2️⃣ Filter for MU / MSFT and select key metrics
# ============================================================
df.columns = [c.strip() for c in df.columns]
target_cols = ["tag", "time", "Trailing P/E", "Forward P/E", "PEG Ratio (5yr expected)"]

filtered = (
    df.loc[df["tag"].isin(["MU", "MSFT"]), target_cols]
    .dropna(how="all", axis=1)
    .reset_index(drop=True)
)

# ============================================================
# 3️⃣ Dash dashboard layout
# ============================================================
app = Dash(__name__)
app.title = "Yahoo Finance Metrics Dashboard"

app.layout = html.Div(
    style={
        "fontFamily": "Arial, sans-serif",
        "margin": "40px",
        "backgroundColor": "#fafafa",
    },
    children=[
        html.H2("📊 Yahoo Finance Key Metrics", style={"textAlign": "center"}),
        html.Hr(),
        html.H4("Trailing P/E, Forward P/E, and PEG Ratio (5yr expected)"),

        # Data Table
        dash_table.DataTable(
            data=filtered.to_dict("records"),
            columns=[{"name": c, "id": c} for c in filtered.columns],
            style_table={"overflowX": "auto"},
            style_header={
                "backgroundColor": "#0078D7",
                "color": "white",
                "fontWeight": "bold",
            },
            style_cell={"padding": "10px", "textAlign": "center"},
        ),

        html.Br(),

        # Trailing P/E Bar Chart
        dcc.Graph(
            figure={
                "data": [
                    {
                        "x": filtered[filtered["tag"] == t]["time"],
                        "y": filtered[filtered["tag"] == t]["Trailing P/E"],
                        "type": "bar",
                        "name": f"{t} Trailing P/E",
                    }
                    for t in filtered["tag"].unique()
                ],
                "layout": {
                    "title": "Trailing P/E Comparison",
                    "barmode": "group",
                    "xaxis": {"title": "Time"},
                    "yaxis": {"title": "P/E Ratio"},
                },
            }
        ),

        # ✅ Forward P/E Bar Chart (new)
        dcc.Graph(
            figure={
                "data": [
                    {
                        "x": filtered[filtered["tag"] == t]["time"],
                        "y": filtered[filtered["tag"] == t]["Forward P/E"],
                        "type": "bar",
                        "name": f"{t} Forward P/E",
                    }
                    for t in filtered["tag"].unique()
                ],
                "layout": {
                    "title": "Forward P/E Comparison",
                    "barmode": "group",
                    "xaxis": {"title": "Time"},
                    "yaxis": {"title": "Forward P/E"},
                },
            }
        ),

        # PEG Ratio Line Chart
        dcc.Graph(
            figure={
                "data": [
                    {
                        "x": filtered[filtered["tag"] == t]["time"],
                        "y": filtered[filtered["tag"] == t]["PEG Ratio (5yr expected)"],
                        "type": "line",
                        "name": f"{t} PEG Ratio",
                    }
                    for t in filtered["tag"].unique()
                ],
                "layout": {
                    "title": "PEG Ratio (5yr expected)",
                    "xaxis": {"title": "Time"},
                    "yaxis": {"title": "PEG"},
                },
            }
        ),
    ],
)

# ============================================================
# 4️⃣ Run locally or on Render
# ============================================================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8050))
    app.run(host="0.0.0.0", port=port, debug=False)
