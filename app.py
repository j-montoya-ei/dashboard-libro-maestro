from pathlib import Path
import unicodedata

import pandas as pd
import plotly.express as px
from dash import Dash, Input, Output, State, ctx, dash_table, dcc, html


DATA_FILE = Path(__file__).parent / "M-FT-1 LIBRO MAESTRO DE DOCUMENTOS v2.xlsx"
ACCENT = "#12355b"
CARD_STYLE = {
    "backgroundColor": "#f7f9fc",
    "border": "1px solid #e5eaf2",
    "borderRadius": "10px",
    "padding": "14px 16px",
}


def clean_text(value: object) -> str:
    text = str(value).strip()
    return "" if text.lower() in {"nan", "nat", "none"} else text


def normalize_header(value: object) -> str:
    text = clean_text(value).replace("�", "O")
    text = unicodedata.normalize("NFKD", text)
    return "".join(char for char in text if not unicodedata.combining(char)).upper()


def load_data() -> pd.DataFrame:
    df = pd.read_excel(DATA_FILE, sheet_name="LIBRO MAESTRO", header=2)
    df.columns = [normalize_header(column) for column in df.columns]
    df = df.loc[:, ~df.columns.str.startswith("UNNAMED")].dropna(how="all").copy()
    for column in df.select_dtypes(include="object").columns:
        df[column] = df[column].map(clean_text)
    for column in [column for column in df.columns if "FECHA" in column]:
        df[column] = pd.to_datetime(df[column], errors="coerce")

    creation, approval, review = "FECHA DE CREACION", "FECHA DE APROBACION", "FECHA DE REVISION"
    df["DIAS HASTA APROBACION"] = (
        (df[approval] - df[creation]).dt.days
        if {creation, approval}.issubset(df.columns)
        else pd.Series(pd.NA, index=df.index, dtype="Float64")
    )
    df["DIAS HASTA REVISION"] = (
        (df[review] - df[creation]).dt.days
        if {creation, review}.issubset(df.columns)
        else pd.Series(pd.NA, index=df.index, dtype="Float64")
    )
    return df


DF = load_data()


def dropdown(column: str, label: str) -> html.Div:
    values = sorted(value for value in DF.get(column, pd.Series(dtype=str)).dropna().unique() if value)
    return html.Div(
        [html.Label(label, className="filter-label"), dcc.Dropdown(
            id=f"filter-{column.lower().replace(' ', '-')}",
            options=[{"label": value, "value": value} for value in values],
            multi=True,
            placeholder=f"Todos los valores de {label.lower()}",
        )],
        className="filter",
    )


def metric(title: str, value_id: str) -> html.Div:
    return html.Div([html.Div(title, className="metric-title"), html.Div(id=value_id, className="metric-value")], style=CARD_STYLE)


app = Dash(__name__, title="Libro Maestro | Centro de control")
server = app.server

app.layout = html.Div(
    [
        html.Div(
            [
                html.Div(
                    [
                        html.Div("LM", className="brand-mark"),
                        html.Div([html.H1("Libro Maestro de Documentos"), html.P("Centro de control documental · Dash + Plotly")], className="header-copy"),
                    ],
                    className="header-content",
                ),
                html.Div("●  Panel activo", className="header-status"),
            ],
            className="header",
        ),
        html.Div(
            [
                html.Div([html.Div("LM", className="sidebar-mark"), html.Div([html.Strong("LIBRO MAESTRO"), html.Small("Gestión documental")])], className="sidebar-brand"),
                html.H3("🎛️ Segmentadores"),
                html.Button("✕ Limpiar selección de gráficos", id="clear-chart-selection", className="clear-button"),
                html.Div(id="chart-selection-label", className="selection-label"),
                dropdown("UNIDAD DE NEGOCIO", "Unidad de negocio"),
                dropdown("MACROPROCESO", "Macroproceso"),
                dropdown("PROCESO", "Proceso"),
                dropdown("ESTADO", "Estado"),
                dropdown("TIPO DE DOCUMENTO", "Tipo de documento"),
                html.Label("Fecha de creación", className="filter-label"),
                dcc.DatePickerRange(id="filter-dates", display_format="YYYY-MM-DD", clearable=True),
            ],
            className="sidebar",
        ),
        dcc.Store(id="chart-selection"),
        html.Main(
            [
                html.H2("📌 Indicadores clave"),
                html.Div([metric("Documentos", "metric-total"), metric("Aprobados", "metric-approved"), metric("No creados", "metric-not-created"), metric("Promedio aprobación", "metric-approval-days"), metric("Promedio revisión", "metric-review-days")], className="metrics"),
                html.Div([dcc.Graph(id="status-chart"), dcc.Graph(id="unit-chart")], className="chart-row"),
                html.Div([dcc.Graph(id="type-chart"), dcc.Graph(id="process-chart")], className="chart-row"),
                html.Div(
                    [
                        html.H2("📋 Detalle del libro maestro"),
                        html.Div(id="result-count"),
                        html.Button("⬇️ Descargar datos filtrados (CSV)", id="download-button", className="download-button"),
                        dcc.Download(id="download-data"),
                        dash_table.DataTable(
                            id="data-table",
                            page_size=15,
                            sort_action="native",
                            filter_action="native",
                            style_table={"overflowX": "auto"},
                            style_cell={"textAlign": "left", "padding": "8px", "minWidth": "130px", "maxWidth": "320px", "whiteSpace": "normal"},
                            style_header={"backgroundColor": ACCENT, "color": "white", "fontWeight": "bold"},
                        ),
                    ],
                    className="detail",
                ),
            ],
            className="content",
        ),
    ],
    className="app-shell",
)


def filter_data(unit, macro, process, status, doc_type, start_date, end_date, chart_selection=None):
    filtered = DF.copy()
    for column, selected in [("UNIDAD DE NEGOCIO", unit), ("MACROPROCESO", macro), ("PROCESO", process), ("ESTADO", status), ("TIPO DE DOCUMENTO", doc_type)]:
        if selected:
            filtered = filtered[filtered[column].isin(selected)]
    if start_date and end_date and "FECHA DE CREACION" in filtered:
        filtered = filtered[filtered["FECHA DE CREACION"].between(pd.Timestamp(start_date), pd.Timestamp(end_date) + pd.Timedelta(days=1))]
    if chart_selection:
        filtered = filtered[filtered[chart_selection["column"]] == chart_selection["value"]]
    return filtered


def chart_or_empty(message: str):
    return {"data": [], "layout": {"annotations": [{"text": message, "showarrow": False}], "xaxis": {"visible": False}, "yaxis": {"visible": False}}}


@app.callback(
    Output("metric-total", "children"), Output("metric-approved", "children"), Output("metric-not-created", "children"),
    Output("metric-approval-days", "children"), Output("metric-review-days", "children"), Output("result-count", "children"),
    Output("status-chart", "figure"), Output("unit-chart", "figure"), Output("type-chart", "figure"), Output("process-chart", "figure"),
    Output("data-table", "data"), Output("data-table", "columns"), Output("chart-selection-label", "children"),
    Input("filter-unidad-de-negocio", "value"), Input("filter-macroproceso", "value"), Input("filter-proceso", "value"),
    Input("filter-estado", "value"), Input("filter-tipo-de-documento", "value"), Input("filter-dates", "start_date"), Input("filter-dates", "end_date"),
    Input("chart-selection", "data"),
)
def update_dashboard(unit, macro, process, status, doc_type, start_date, end_date, chart_selection):
    filtered = filter_data(unit, macro, process, status, doc_type, start_date, end_date, chart_selection)
    approved = int((filtered.get("ESTADO", pd.Series(index=filtered.index)) == "Aprobado").sum())
    not_created = int((filtered.get("ESTADO", pd.Series(index=filtered.index)) == "No creado").sum())
    approval_days, review_days = filtered["DIAS HASTA APROBACION"].mean(), filtered["DIAS HASTA REVISION"].mean()

    if filtered.empty:
        empty = chart_or_empty("No hay datos para los filtros seleccionados.")
        figures = [empty] * 4
    else:
        status_data = filtered["ESTADO"].replace("", "Sin estado").value_counts().rename_axis("ESTADO").reset_index(name="CANTIDAD")
        unit_data = filtered["UNIDAD DE NEGOCIO"].replace("", "Sin unidad").value_counts().rename_axis("UNIDAD").reset_index(name="CANTIDAD")
        type_data = filtered["TIPO DE DOCUMENTO"].replace("", "Sin tipo").value_counts().head(12).sort_values()
        process_data = filtered.groupby("PROCESO")["DIAS HASTA APROBACION"].mean().dropna().nlargest(10).sort_values()
        figures = [
            px.pie(status_data, names="ESTADO", values="CANTIDAD", hole=.55, title="Documentos por estado"),
            px.bar(unit_data, x="UNIDAD", y="CANTIDAD", color="UNIDAD", text="CANTIDAD", title="Documentos por unidad de negocio"),
            px.bar(type_data, x=type_data.values, y=type_data.index, orientation="h", text=type_data.values, title="Tipos de documento"),
            px.bar(process_data, x=process_data.values, y=process_data.index, orientation="h", text_auto=".0f", color=process_data.values, title="Procesos con mayor tiempo de ciclo") if not process_data.empty else chart_or_empty("No hay fechas suficientes para calcular tiempos."),
        ]
    for figure in figures:
        figure.update_layout(template="plotly_white", margin={"t": 55, "b": 25, "l": 25, "r": 25})

    display = filtered.copy()
    for column in display.columns:
        if "FECHA" in column:
            display[column] = display[column].dt.strftime("%Y-%m-%d")
    display = display.fillna("")
    return (
        f"{len(filtered):,}", f"{approved:,}", f"{not_created:,}",
        f"{approval_days:.0f} días" if pd.notna(approval_days) else "N/D",
        f"{review_days:.0f} días" if pd.notna(review_days) else "N/D",
        f"{len(filtered):,} registros encontrados", figures[0], figures[1], figures[2], figures[3],
        display.to_dict("records"), [{"name": column, "id": column} for column in display.columns],
        f"🔎 Selección activa: {chart_selection['value']}" if chart_selection else "💡 Haz clic en una barra o segmento para filtrar todo el dashboard.",
    )


@app.callback(
    Output("chart-selection", "data"),
    Input("status-chart", "clickData"), Input("unit-chart", "clickData"),
    Input("type-chart", "clickData"), Input("process-chart", "clickData"),
    Input("clear-chart-selection", "n_clicks"),
)
def update_chart_selection(status_click, unit_click, type_click, process_click, clear_clicks):
    if ctx.triggered_id == "clear-chart-selection":
        return None
    click_map = [
        ("status-chart", status_click, "ESTADO", "label"),
        ("unit-chart", unit_click, "UNIDAD DE NEGOCIO", "x"),
        ("type-chart", type_click, "TIPO DE DOCUMENTO", "y"),
        ("process-chart", process_click, "PROCESO", "y"),
    ]
    for chart_id, click_data, column, axis in click_map:
        if ctx.triggered_id == chart_id and click_data and click_data.get("points"):
            return {"column": column, "value": click_data["points"][0].get(axis)}
    return None


@app.callback(Output("download-data", "data"), Input("download-button", "n_clicks"), State("filter-unidad-de-negocio", "value"), State("filter-macroproceso", "value"), State("filter-proceso", "value"), State("filter-estado", "value"), State("filter-tipo-de-documento", "value"), State("filter-dates", "start_date"), State("filter-dates", "end_date"), State("chart-selection", "data"), prevent_initial_call=True)
def download_data(_, unit, macro, process, status, doc_type, start_date, end_date, chart_selection):
    filtered = filter_data(unit, macro, process, status, doc_type, start_date, end_date, chart_selection)
    return dcc.send_data_frame(filtered.to_csv, "libro_maestro_filtrado.csv", index=False, encoding="utf-8-sig")


if __name__ == "__main__":
    app.run(debug=False, host="127.0.0.1", port=8050)
