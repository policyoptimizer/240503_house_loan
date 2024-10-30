# 날짜 도출 서식 구현하는 것은 포기함
# 실질적으로 필요도 없어보임
# 관리하는 엑셀파일을 전처리해서 관리하기 쉽게 도출해주는 역할을 함

import dash
from dash import dcc, html
from dash.dependencies import Input, Output, State
import pandas as pd
from datetime import datetime
import io
import base64
import dash.dash_table as dash_table

# 오늘 날짜 설정
TODAY = datetime(2024, 10, 30)

#app = dash.Dash(__name__)

app.layout = html.Div([
    html.H1('대출 서류 제출 일정 관리'),
    dcc.Upload(
        id='upload-data',
        children=html.Div([
            '엑셀 파일을 여기로 드래그하거나 ',
            html.A('클릭하여 업로드')
        ]),
        style={
            'width': '60%',
            'height': '60px',
            'lineHeight': '60px',
            'borderWidth': '1px',
            'borderStyle': 'dashed',
            'borderRadius': '5px',
            'textAlign': 'center',
            'margin': '10px'
        },
        multiple=False,
        accept='.xlsx'
    ),
    html.Button('임박한 대상자 추출', id='extract-button', n_clicks=0, style={'display': 'none'}),
    html.Div(id='output-data-upload'),
    html.Div(id='notification-result'),
    html.A(
        '추출된 데이터 다운로드',
        id='download-link',
        download="imminent_submissions.xlsx",
        href="",
        target="_blank",
        style={'display': 'none'}
    )
])

def parse_contents(contents, filename):
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    # 엑셀 파일 읽기
    df_list = []
    xls = pd.ExcelFile(io.BytesIO(decoded))
    sheet_names = xls.sheet_names

    for sheet_name in sheet_names:
        df_sheet = pd.read_excel(xls, sheet_name=sheet_name)
        df_list.append(df_sheet)

    # 시트들을 하나의 데이터프레임으로 통합
    df = pd.concat(df_list, ignore_index=True)
    return df

@app.callback(
    Output('output-data-upload', 'children'),
    Output('extract-button', 'style'),
    Input('upload-data', 'contents'),
    State('upload-data', 'filename')
)
def update_output(contents, filename):
    if contents is not None:
        df = parse_contents(contents, filename)

        # 업로드된 데이터 표시
        children = html.Div([
            html.H5(f"업로드된 파일: {filename}"),
            html.H6("전체 데이터 미리보기:"),
            dash_table.DataTable(
                data=df.head(10).to_dict('records'),
                columns=[{'name': i, 'id': i} for i in df.columns],
                page_size=10
            ),
            html.Br()
        ])

        # 추출 버튼 보이기
        extract_button_style = {'display': 'block', 'margin': '10px'}

        return children, extract_button_style
    else:
        return '', {'display': 'none'}

@app.callback(
    Output('notification-result', 'children'),
    Output('download-link', 'href'),
    Output('download-link', 'style'),
    Input('extract-button', 'n_clicks'),
    State('upload-data', 'contents'),
    State('upload-data', 'filename')
)
def extract_imminent_submissions(n_clicks, contents, filename):
    if n_clicks > 0 and contents is not None:
        df = parse_contents(contents, filename)

        # 날짜 칼럼 파싱
        date_columns = [
            '유주택자 증빙서류(등기부등본_기존)',
            '구매증빙서류(등기부등본_신규)',
            '구매증빙서류(주민등록등본_신규)',
            '전세증빙서류(주민등록등본)'
        ]

        for col in date_columns:
            df[col] = pd.to_datetime(df[col], errors='coerce', format='%Y-%m-%d')

        # 임박한 날짜 기준 필터링 (오늘부터 2개월 이내)
        two_months_later = TODAY + pd.DateOffset(months=2)

        # 각 날짜 칼럼에 대해 필터링하고 대상자 추출
        imminent_df_list = []
        for col in date_columns:
            temp_df = df[
                (df[col] >= TODAY) &
                (df[col] <= two_months_later)
            ]
            if not temp_df.empty:
                temp_df = temp_df.copy()
                temp_df['임박한 서류'] = col
                temp_df['제출기한'] = temp_df[col].dt.strftime('%Y-%m-%d')
                imminent_df_list.append(temp_df)

        if imminent_df_list:
            imminent_df = pd.concat(imminent_df_list, ignore_index=True)
            # 중복 제거
            imminent_df = imminent_df.drop_duplicates(subset=['사번', '임박한 서류'])

            # 필요한 칼럼만 선택
            display_columns = ['성명', '사번', '직위', '임박한 서류', '제출기한', '서류제출완료여부']
            imminent_df_display = imminent_df[display_columns]

            # 결과 표시
            children = html.Div([
                html.H5("임박한 대상자 목록:"),
                dash_table.DataTable(
                    data=imminent_df_display.to_dict('records'),
                    columns=[{'name': i, 'id': i} for i in imminent_df_display.columns],
                    page_size=20
                ),
                html.Br()
            ])

            # 결과를 엑셀 파일로 저장
            output = io.BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            imminent_df_display.to_excel(writer, index=False)
            writer.save()
            processed_data = output.getvalue()

            # 다운로드 링크 생성
            href_data = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + base64.b64encode(processed_data).decode()

            download_link_style = {'display': 'block', 'margin': '10px'}

            return children, href_data, download_link_style
        else:
            return html.Div("임박한 대상자가 없습니다."), '', {'display': 'none'}
    else:
        return '', '', {'display': 'none'}

# Dataiku에서 앱을 실행할 때는 아래 부분을 주석 처리해야 합니다.
# if __name__ == '__main__':
#     app.run_server(debug=True)
