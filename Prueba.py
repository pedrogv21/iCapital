# -*- coding: utf-8 -*-
"""
Created on Mon Jul 22 16:20:29 2024

@author: USUARIO
"""
import dash
from dash import dcc, html
from dash.dependencies import Input, Output, State
import plotly.express as px
import pandas as pd
from dash import dash_table

# Paths to the Excel files
main_excel_file_path = r'C:\Users\USUARIO\Seguimiento Unicorn\Unicorn Report.xlsx'
mapping_excel_file_path = r'C:\Users\USUARIO\Seguimiento Unicorn\column mapping.xlsx'

# Read the main Excel file
try:
    df = pd.read_excel(main_excel_file_path)
    print("Archivo de Excel leído correctamente.")
    print(f"Columnas encontradas en el DataFrame: {df.columns.tolist()}")
except Exception as e:
    print(f"Error leyendo el archivo de Excel: {e}")
    df = pd.DataFrame()  # Create an empty DataFrame in case of an error

# Read the mapping Excel file
try:
    mapping_df = pd.read_excel(mapping_excel_file_path)
    value_mapping = dict(zip(mapping_df['Fund Name'], mapping_df['Nombre Uso Interno']))
    print("Archivo de mapeo leído correctamente.")
    print(f"Mapeo de valores: {value_mapping}")
except Exception as e:
    print(f"Error leyendo el archivo de mapeo: {e}")
    value_mapping = {}

# Apply the value mapping to the 'Fund Name' column
if 'Fund Name' in df.columns and value_mapping:
    df['Fund Name'] = df['Fund Name'].map(value_mapping).fillna(df['Fund Name'])

# Print the columns to check if 'User  Firm' exists
print(f"Columnas después del mapeo: {df.columns.tolist()}")

# Check that the expected columns exist in the DataFrame after renaming
required_columns = {'Sales Manager', 'Close Date', 'Investment Amount', 'Fund Name', 'Review Status', 'User  Firm'}
missing_columns = required_columns - set(df.columns)

if missing_columns:
    print(f"Error: El archivo de Excel no contiene las columnas necesarias después de renombrar. Columnas faltantes: {missing_columns}")
    df = pd.DataFrame()  # Create an empty DataFrame if columns are missing
else:
    # Replace empty values in 'Investment Amount' with 0
    df['Investment Amount'] = df['Investment Amount'].fillna(0)
    # Replace zero values in 'Close Date' with NaT
    df['Close Date'] = pd.to_datetime(df['Close Date'], errors='coerce')
    # Clean up the 'Sales Manager' column
    df['Sales Manager'] = df['Sales Manager'].str.replace(r'@unicornsp.com#0', '', regex=True).str.replace('.', ' ')

# Initialize the Dash app
app = dash.Dash(__name__)
app.config.suppress_callback_exceptions = True

# App layout with tabs
if not df.empty:
    app.layout = html.Div(children=[
        html.H1(children='Sales Dashboard'),
        dcc.Tabs(id="tabs-example", value='tab-1', children=[
            dcc.Tab(label='Tab 1: Overview', value='tab-1'),
            dcc.Tab(label='Tab 2: Details', value='tab-2'),
        ]),
        html.Div(id='tabs-content-example')
    ])
else:
    app.layout = html.Div(children=[
        html.H1(children='Sales Dashboard'),
        html.Div(children='Error: El archivo de Excel no contiene las columnas necesarias.')
    ])

# Callback to update the content based on the selected tab
@app.callback(
    Output('tabs-content-example', 'children'),
    [Input('tabs-example', 'value')]
)
def render_content(tab):
    if tab == 'tab-1':
        return html.Div([
            dcc.Dropdown(
                id='sales-manager-dropdown',
                options=[{'label': 'Todos', 'value': 'Todos'}] +
                        [{'label': manager, 'value': manager} for manager in df['Sales Manager'].unique()],
                value=['Todos'],
                multi=True
            ),
            
            dcc.Dropdown(
                id='fund-name-dropdown',
                options=[{'label': 'Todos', 'value': 'Todos'}] +
                        [{'label': fund, 'value': fund} for fund in df['Fund Name'].unique()],
                value=['Todos'],
                multi=True
            ),

            dcc.Graph(
                id='investment-graph'
            ),
            
            html.Div(id='click-data')
        ])
    elif tab == 'tab-2':
        return html.Div([
            dcc.Dropdown(
                id='review-status-dropdown',
                options=[{'label': 'Todos', 'value': 'Todos'}] +
                        [{'label': status, 'value': status} for status in df['Review Status'].unique()],
                value=['Todos'],
                multi=True
            ),
            dcc.Dropdown(
                id='sales-manager-dropdown-tab2',
                options=[{'label': 'Todos', 'value': 'Todos'}] +
                        [{'label': manager, 'value': manager} for manager in df['Sales Manager'].unique()],
                value=['Todos'],
                multi=True
            ),
            dcc.Dropdown(
                id='user-firm-dropdown',
                options=[{'label': 'Todos', 'value': 'Todos'}] +
                        [{'label': firm, 'value': firm} for firm in df['User  Firm'].unique()],
                value=['Todos'],
                multi=True
            ),
            html.Div(id='table-container')
        ])

# Callback to update the graph in Tab 1
@app.callback(
    Output('investment-graph', 'figure'),
    [Input('sales-manager-dropdown', 'value'),
     Input('fund-name-dropdown', 'value')]
)
def update_graph(selected_managers, selected_funds):
    print(f"Selected managers: {selected_managers}, Selected funds: {selected_funds}")
    if df.empty or not selected_managers or not selected_funds:
        print("DataFrame está vacío o no se han seleccionado los filtros necesarios.")
        return {}  # Return an empty figure if no data

    try:
        filtered_df = df.copy()
        
        if 'Todos' not in selected_managers:
            filtered_df = filtered_df[filtered_df['Sales Manager'].isin(selected_managers)]
        
        if 'Todos' not in selected_funds:
            filtered_df = filtered_df[filtered_df['Fund Name'].isin(selected_funds)]

        print(f"Filtrado DataFrame:\n{filtered_df}")

        filtered_df['Close Date'] = pd.to_datetime(filtered_df['Close Date'], errors='coerce')
        
        # Create a separate column for NaT values
        nat_df = filtered_df[filtered_df['Close Date'].isna()]
        nat_sum = nat_df['Investment Amount'].sum()

        # Exclude NaT values from the main DataFrame
        filtered_df = filtered_df.dropna(subset=['Close Date'])

        # Group by month and sum 'Investment Amount'
        filtered_df['Close Date'] = filtered_df['Close Date'].dt.to_period('M')
        grouped_df = filtered_df.groupby('Close Date')['Investment Amount'].sum().reset_index()
        grouped_df['Close Date'] = grouped_df['Close Date'].dt.to_timestamp()

        # Create a complete range of months
        if not grouped_df.empty:
            full_range = pd.date_range(start=grouped_df['Close Date'].min(), end=grouped_df['Close Date'].max(), freq='MS')
            full_df = pd.DataFrame(full_range, columns=['Close Date'])

            # Merge the complete range with the aggregated data
            full_df = full_df.merge(grouped_df, on='Close Date', how='left').fillna({'Investment Amount': 0})
        else:
            full_df = pd.DataFrame(columns=['Close Date', 'Investment Amount'])

        # Add the NaT row to the DataFrame
        if nat_sum > 0:
            nat_row = pd.DataFrame({'Close Date': ['N/A'], 'Investment Amount': [nat_sum]})
            full_df = pd.concat([full_df, nat_row], ignore_index=True)

        # Convert dates to text format
        full_df['Close Date'] = full_df['Close Date'].apply(lambda x: x.strftime('%B %Y') if x != 'N/A' else 'N/A')

        print(f"DataFrame con rango completo de meses:\n{full_df}")

        fig = px.bar(full_df, x='Close Date', y='Investment Amount', title='Investment Amounts')
        fig.update_traces(texttemplate='%{y}', textposition='outside')
        fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')

        # Ensure the x-axis shows all available months and years
        fig.update_xaxes(type='category')
        return fig
    except Exception as e:
        print(f"Error en el callback: {e}")
        return {}

# Callback to handle clicking on the bars in the graph
@app.callback(
    Output('click-data', 'children'),
    [Input('investment-graph', 'clickData')],
    [State('sales-manager-dropdown', 'value'),
     State('fund-name-dropdown', 'value')]
)
def display_click_data(clickData, selected_managers, selected_funds):
    if clickData is None:
        return "Haz clic en una barra para ver más información."
    
    # Extract click information
    point_data = clickData['points'][0]
    x_value = point_data['x']
    
    # Filter the original DataFrame based on the selected filters
    filtered_df = df.copy()
    
    if 'Todos' not in selected_managers:
        filtered_df = filtered_df[filtered_df['Sales Manager'].isin(selected_managers)]
    
    if 'Todos' not in selected_funds:
        filtered_df = filtered_df[filtered_df['Fund Name'].isin(selected_funds)]
    
    if x_value == 'N/A':
        filtered_df = filtered_df[filtered_df['Close Date'].isna()]
    else:
        date_filter = pd.to_datetime(x_value, format='%B %Y', errors='coerce')
        filtered_df = filtered_df[filtered_df['Close Date'].dt.to_period('M') == date_filter.to_period('M')]
    
    # Filter the columns to be displayed
    available_columns = filtered_df.columns
    required_display_columns = ['Sales Manager', 'Advisor Firm Name', 'Name', 'Advisor Name', 'User Name', 'User  Firm', 'Fund Name', 'Investment Amount']
    display_columns = [col for col in required_display_columns if col in available_columns]
    filtered_df = filtered_df[display_columns]

    # Format the 'Investment Amount' column for display
    if 'Investment Amount' in filtered_df.columns:
        filtered_df['Investment Amount'] = filtered_df['Investment Amount'].apply(lambda x: "{:,.2f}".format(x))
    
    # Create a table with the filtered data
    return html.Div([
        html.H4(f"Datos para {x_value}"),
        dcc.Graph(
            id='table-graph',
            figure=px.bar(filtered_df, x='Sales Manager', y='Investment Amount', color='Fund Name', barmode='group')
        ),
        html.Table([
            html.Thead(
                html.Tr([html.Th(col) for col in filtered_df.columns])
            ),
            html.Tbody([
                html.Tr([
                    html.Td(filtered_df.iloc[i][col]) for col in filtered_df.columns
                ]) for i in range(len(filtered_df))
            ])
        ])
    ])

# Callback to update the table in Tab 2
@app.callback(
    Output('table-container', 'children'),
    [Input('review-status-dropdown', 'value'),
     Input('sales-manager-dropdown-tab2', 'value'),
     Input('user-firm-dropdown', 'value')]
)
def update_table(selected_status, selected_managers, selected_user_firms):
    if not selected_status or not selected_managers or not selected_user_firms:
        return "Please select a review status, sales manager, and user firm to display the table."

    filtered_df = df.copy()
    
    if 'Todos' not in selected_status:
        filtered_df = filtered_df[filtered_df['Review Status'].isin(selected_status)]
    
    if 'Todos' not in selected_managers:
        filtered_df = filtered_df[filtered_df['Sales Manager'].isin(selected_managers)]
    
    if 'Todos' not in selected_user_firms:
        filtered_df = filtered_df[filtered_df['User  Firm'].isin(selected_user_firms)]
    
    # Drop the specified columns
    columns_to_drop = ['Sales Referral Code ID', 'Unicorn Region', 'Advisor Email', 'User Email', 'Portal', 'Fund Jurisdiction', 'Canceled', 'Available']
    filtered_df.drop(columns=columns_to_drop, inplace=True, errors='ignore')
    
    if filtered_df.empty:
        return "No data available for the selected filters."

    # Format 'Investment Amount' as a number with thousands separators
    if 'Investment Amount' in filtered_df.columns:
        filtered_df['Investment Amount'] = filtered_df['Investment Amount'].apply(lambda x: "{:,.2f}".format(x))

    # Format 'Close Date' and 'Last Status Update' to short date format
    date_columns = ['Close Date', 'Last Status Update']
    for col in date_columns:
        if col in filtered_df.columns:
            filtered_df[col] = pd.to_datetime(filtered_df[col], errors='coerce').dt.strftime('%d-%m-%Y')

    return html.Div([
        dash_table.DataTable(
            columns=[{"name": col, "id": col} for col in filtered_df.columns],
            data=filtered_df.to_dict('records'),
            filter_action='native',
            sort_action='native',
            page_action='none',  # Display all rows without pagination
            style_table={'overflowX': 'auto', 'maxHeight': '600px', 'overflowY': 'auto'},  # Enable vertical scroll if needed
            export_format='xlsx'
        )
    ])

# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)
