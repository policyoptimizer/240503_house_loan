import dash
from dash import dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
import pandas as pd
from datetime import datetime, timedelta
import base64
import io

# Initialize the Dash app
# app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server  # For deployment purposes

# Define the layout of the app
app.layout = dbc.Container([
    dbc.Row([
        dbc.Col(html.H1("주택자금 증빙서류 관리"), className="mb-4")
    ]),
    dbc.Row([
        dbc.Col([
            # File upload component
            dcc.Upload(
                id='upload-data',
                children=html.Div([
                    '엑셀 파일을 드래그하거나 ',
                    html.A('여기를 클릭하여 업로드')
                ]),
                style={
                    'width': '100%',
                    'height': '60px',
                    'lineHeight': '60px',
                    'borderWidth': '1px',
                    'borderStyle': 'dashed',
                    'borderRadius': '5px',
                    'textAlign': 'center',
                    'margin-bottom': '10px'
                },
                multiple=False
            ),
            # Button to process the file
            dbc.Button('파일 처리하기', id='process-button', color='primary', disabled=True),
            # Button to download the processed file
            dbc.Button('처리된 파일 다운로드', id='download-button', color='success', disabled=True, style={'margin-left': '10px'}),
            # Output messages
            html.Div(id='output-message', style={'margin-top': '20px'}),
            # Hidden components to store data
            dcc.Download(id='download-dataframe-csv'),
            dcc.Store(id='uploaded-data'),
            dcc.Store(id='processed-data')
        ])
    ])
])

# Function to parse dates in standard format
def parse_date_standard(date_str):
    if pd.isna(date_str):
        return date_str
    try:
        for fmt in ("%Y년 %m월", "%Y-%m-%d", "%Y-%m", "%b-%d"):
            try:
                dt = datetime.strptime(str(date_str), fmt)
                return dt
            except ValueError:
                continue
        return pd.to_datetime(date_str, errors='coerce')
    except Exception:
        return pd.NaT

# Function to calculate dates based on conditions
def calculate_date(base_date, offset_days, condition=True):
    if not condition or pd.isna(base_date):
        return "미대상"
    result_date = base_date + timedelta(days=offset_days)
    return result_date.strftime("%Y년 %m월 %d일")

# Callback to store the uploaded file and enable the process button
@app.callback(
    Output('uploaded-data', 'data'),
    Output('process-button', 'disabled'),
    Input('upload-data', 'contents'),
    State('upload-data', 'filename')
)
def store_uploaded_data(contents, filename):
    if contents is not None:
        return {'contents': contents, 'filename': filename}, False
    else:
        return dash.no_update, True

# Callback to process the data when the process button is clicked
@app.callback(
    Output('processed-data', 'data'),
    Output('output-message', 'children'),
    Output('download-button', 'disabled'),
    Input('process-button', 'n_clicks'),
    State('uploaded-data', 'data')
)
def process_data(n_clicks, uploaded_data):
    if n_clicks is None or uploaded_data is None:
        return dash.no_update, dash.no_update, True
    else:
        contents = uploaded_data['contents']
        filename = uploaded_data['filename']
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)
        try:
            # Read the Excel or CSV file
            if 'xls' in filename:
                df = pd.read_excel(io.BytesIO(decoded))
            elif 'csv' in filename:
                df = pd.read_csv(io.StringIO(decoded.decode('utf-8')))
            else:
                return dash.no_update, html.Div(['지원되지 않는 파일 형식입니다.']), True

            # Parse date columns
            date_columns = ['대출신청월', '입주예정일']
            for col in date_columns:
                df[col] = df[col].apply(parse_date_standard)

            # Process each row according to specified logic
            def process_row(row):
                # 유주택자 증빙서류(등기부등본_기존)
                if row.get('유주택/무주택(서약서대상여부)', '') == '대상':
                    row['유주택자 증빙서류(등기부등본_기존)'] = calculate_date(row['대출신청월'], -30)
                else:
                    row['유주택자 증빙서류(등기부등본_기존)'] = '미대상'

                # 구매증빙서류(등기부등본_신규)
                if row.get('구매구분', '') in ['구매', '분양']:
                    row['구매증빙서류(등기부등본_신규)'] = calculate_date(row['입주예정일'], 90)
                else:
                    row['구매증빙서류(등기부등본_신규)'] = '미대상'

                # 구매증빙서류(주민등록등본_신규)
                if row.get('구매구분', '') in ['구매', '분양']:
                    row['구매증빙서류(주민등록등본_신규)'] = calculate_date(row['입주예정일'], 90)
                else:
                    row['구매증빙서류(주민등록등본_신규)'] = '미대상'

                # 전세증빙서류(주민등록등본)
                if row.get('구매구분', '') == '전세':
                    row['전세증빙서류(주민등록등본)'] = calculate_date(row['입주예정일'], 90)
                else:
                    row['전세증빙서류(주민등록등본)'] = '미대상'

                return row

            df = df.apply(process_row, axis=1)

            # Calculate 서류제출완료여부
            def check_documents_status(row):
                documents = [
                    '유주택자 증빙서류(등기부등본_기존)',
                    '구매증빙서류(등기부등본_신규)',
                    '구매증빙서류(주민등록등본_신규)',
                    '전세증빙서류(주민등록등본)'
                ]
                counts = [1 if row.get(col, '') != '미대상' else 0 for col in documents]
                if sum(counts) == len(documents):
                    return "완료"
                elif sum(counts) == 0:
                    return "미진행중"
                else:
                    return "진행중"

            df['서류제출완료여부'] = df.apply(check_documents_status, axis=1)

            # Prepare the CSV for download
            processed_csv = df.to_csv(index=False, encoding='utf-8-sig')
            processed_csv_encoded = base64.b64encode(processed_csv.encode()).decode()

            return {'filename': 'processed_'+filename, 'content': processed_csv_encoded}, html.Div(['파일 처리가 완료되었습니다. 아래에서 다운로드하세요.']), False

        except Exception as e:
            return dash.no_update, html.Div(['파일을 처리하는 중 오류가 발생했습니다: {}'.format(str(e))]), True

# Callback to enable the download button and provide the file for download
@app.callback(
    Output('download-dataframe-csv', 'data'),
    Input('download-button', 'n_clicks'),
    State('processed-data', 'data'),
    prevent_initial_call=True
)
def download_processed_data(n_clicks, processed_data):
    if n_clicks and processed_data:
        decoded = base64.b64decode(processed_data['content'])
        return dict(content=decoded, filename=processed_data['filename'])
    else:
        return dash.no_update

# Run the app
if __name__ == '__main__':
    # Due to internal security policies, the app instance is commented out.
    # app.run_server(debug=True)
    pass
