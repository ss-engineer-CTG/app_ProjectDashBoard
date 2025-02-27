"""
プロジェクト管理ダッシュボードのメインアプリケーション
"""

import logging
import dash
from dash import html, dcc
import os

from ProjectDashBoard.config import COLORS, STYLES, HTML_TEMPLATE
from ProjectDashBoard.callbacks import register_callbacks

# ログディレクトリの作成
log_dir = 'logs'
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

# ロギングの設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    filename=os.path.join(log_dir, 'dashboard.log')
)
logger = logging.getLogger(__name__)

# Dashアプリケーションの初期化
app = dash.Dash(__name__)
app.index_string = HTML_TEMPLATE

# アプリケーションのレイアウト
app.layout = html.Div([
    # 更新ボタン（非表示）
    html.Button('更新', id='update-button', n_clicks=0, style={'display': 'none'}),
    
    # ダミー出力（コールバック用）
    html.Div(id='dummy-output', style={'display': 'none'}),
    
    # 通知コンテナ
    html.Div(id='notification-container'),
    
    # ヘッダー
    html.Div([
        html.Div([
            html.H1('プロジェクト進捗ダッシュボード', 
                   style={'color': COLORS['text']['primary']}),
            html.P(id='update-time',
                  style={'color': COLORS['text']['secondary']})
        ], style=STYLES['container'])
    ], style={'backgroundColor': COLORS['surface'],
              'borderBottom': '1px solid rgba(255,255,255,0.1)',
              'boxShadow': '0 2px 4px rgba(0,0,0,0.2)'}),
    
    # メインコンテンツ
    html.Div([
        # サマリーカード
        html.Div([
            html.Div([
                html.Div([
                    html.H3('総プロジェクト数', style={'color': COLORS['text']['secondary']}),
                    html.H2(id='total-projects', style={'color': COLORS['text']['primary']})
                ], style={**STYLES['card'], 'width': '23%'}),
                html.Div([
                    html.H3('進行中', style={'color': COLORS['text']['secondary']}),
                    html.H2(id='active-projects', style={'color': COLORS['status']['info']})
                ], style={**STYLES['card'], 'width': '23%'}),
                html.Div([
                    html.H3('遅延あり', style={'color': COLORS['text']['secondary']}),
                    html.H2(id='delayed-projects', style={'color': COLORS['status']['danger']})
                ], style={**STYLES['card'], 'width': '23%'}),
                html.Div([
                    html.H3('今月のマイルストーン', style={'color': COLORS['text']['secondary']}),
                    html.H2(id='milestone-projects', style={'color': COLORS['status']['warning']})
                ], style={**STYLES['card'], 'width': '23%'})
            ], style={'display': 'flex', 'justifyContent': 'space-between', 'marginBottom': '20px'})
        ]),
        
        # プロジェクト一覧
        html.Div([
            html.H2('プロジェクト一覧', style=STYLES['header']),
            html.Div(
                html.Div(id='project-table'),
                style={
                    'overflowX': 'auto',  # 横スクロールを有効化
                    'width': '100%',
                    'maxHeight': '600px',  # 高さの制限も設定するとより使いやすくなる
                    'overflowY': 'auto'   # 縦スクロールも有効化
                }
            )
        ], style=STYLES['card']),
        
        # グラフセクション
        html.Div([
            html.Div([
                html.Div([
                    html.H2('進捗状況分布', style=STYLES['header']),
                    dcc.Graph(id='progress-distribution')
                ], style={**STYLES['card'], 'width': '48%'}),
                html.Div([
                    html.H2('プロジェクト期間分布', style=STYLES['header']),
                    dcc.Graph(id='duration-distribution')
                ], style={**STYLES['card'], 'width': '48%'})
            ], style={'display': 'flex', 'justifyContent': 'space-between'})
        ])
    ], style={'backgroundColor': COLORS['background'], 
              'padding': '20px', 
              'minHeight': '100vh'})
], style={'backgroundColor': COLORS['background']})

# コールバックの登録
register_callbacks(app)

# アプリケーション起動
if __name__ == '__main__':
    logger.info("Starting dashboard application")
    app.run_server(debug=True)