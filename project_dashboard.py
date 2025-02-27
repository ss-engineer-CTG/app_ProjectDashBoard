import pandas as pd
import plotly.graph_objects as go
import dash
from dash import html, dcc
from dash.dependencies import Input, Output, State, ALL, MATCH
from dash.exceptions import PreventUpdate
import datetime
import os
import platform
import subprocess
import logging
from pathlib import Path
from typing import Optional, Union, Dict, Any, List, Tuple

# ロギングの設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    filename='dashboard.log'
)
logger = logging.getLogger(__name__)

app = dash.Dash(__name__)

# カラー定義
COLORS = {
    'background': '#1a1a1a',
    'surface': '#2d2d2d',
    'text': {
        'primary': '#ffffff',
        'secondary': '#b3b3b3',
        'accent': '#60cdff'
    },
    'status': {
        'success': '#50ff96',
        'warning': '#ffeb45',
        'danger': '#ff5f5f',
        'info': '#60cdff',
        'neutral': '#c8c8c8'
    },
    'chart': {
        'primary': ['#60cdff', '#50ff96', '#ffeb45', '#ff5f5f', '#ff60d3', '#d160ff'],
        'background': 'rgba(45,45,45,0.9)'
    }
}

# スタイル定義
STYLES = {
    'card': {
        'backgroundColor': COLORS['surface'],
        'padding': '20px',
        'borderRadius': '10px',
        'boxShadow': '0 4px 6px rgba(0,0,0,0.3)',
        'marginBottom': '20px',
        'border': '1px solid rgba(255,255,255,0.1)'
    },
    'header': {
        'color': COLORS['text']['primary'],
        'marginBottom': '15px',
        'fontWeight': 'bold'
    },
    'container': {
        'maxWidth': '1200px',
        'margin': '0 auto',
        'padding': '20px'
    },
    'progressBar': {
        'container': {
            'width': '100%',
            'backgroundColor': 'rgba(255,255,255,0.1)',
            'borderRadius': '4px',
            'overflow': 'hidden',
            'height': '20px',
            'position': 'relative'
        },
        'bar': {
            'height': '100%',
            'transition': 'width 0.3s ease-in-out',
            'borderRadius': '4px'
        },
        'text': {
            'position': 'absolute',
            'top': '0',
            'left': '0',
            'width': '100%',
            'height': '100%',
            'display': 'flex',
            'alignItems': 'center',
            'justifyContent': 'center',
            'color': 'white',
            'fontSize': '12px',
            'textShadow': '1px 1px 2px rgba(0,0,0,0.5)',
            'zIndex': '1'
        }
    },
    'linkButton': {
        'backgroundColor': COLORS['surface'],
        'color': COLORS['text']['accent'],
        'padding': '6px 12px',
        'border': f'1px solid {COLORS["text"]["accent"]}',
        'borderRadius': '4px',
        'textDecoration': 'none',
        'fontSize': '12px',
        'margin': '0 4px',
        'display': 'inline-block',
        'cursor': 'pointer',
        'transition': 'all 0.3s ease'
    },
    'notification': {
        'position': 'fixed',
        'bottom': '20px',
        'right': '20px',
        'padding': '10px 20px',
        'borderRadius': '4px',
        'color': 'white',
        'zIndex': '1000',
        'transition': 'opacity 0.3s ease-in-out',
        'success': {
            'backgroundColor': COLORS['status']['success']
        },
        'error': {
            'backgroundColor': COLORS['status']['danger']
        }
    }
}

# グラフ共通レイアウト
GRAPH_LAYOUT = {
    'paper_bgcolor': 'rgba(0,0,0,0)',
    'plot_bgcolor': 'rgba(0,0,0,0)',
    'font': {'color': COLORS['text']['primary']},
    'xaxis': {
        'gridcolor': 'rgba(255,255,255,0.1)',
        'zerolinecolor': 'rgba(255,255,255,0.1)'
    },
    'yaxis': {
        'gridcolor': 'rgba(255,255,255,0.1)',
        'zerolinecolor': 'rgba(255,255,255,0.1)'
    }
}

# スタイルシートの設定
app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <style>
            .link-button:hover {
                background-color: ''' + COLORS['text']['accent'] + ''' !important;
                color: ''' + COLORS['surface'] + ''' !important;
            }
        </style>
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>
'''

# ユーティリティ関数
def validate_file_path(path: Optional[str], allow_directories: bool = True) -> Optional[str]:
    """
    ファイルパスの検証と正規化を行う
    
    Args:
        path: 検証するファイルパス
        allow_directories: ディレクトリを許可するかどうか
        
    Returns:
        検証済みの正規化されたパス、または無効な場合はNone
    """
    try:
        if not path or pd.isna(path):
            logger.warning("Empty or NaN path provided")
            return None
            
        # パスの正規化前にデバッグ情報
        logger.info(f"Original path: {path}")
        
        # 文字列型の確認と変換
        if not isinstance(path, str):
            path = str(path)
        
        # パスの正規化
        normalized_path = str(Path(path).resolve())
        logger.info(f"Normalized path: {normalized_path}")
        
        # ディレクトリの場合は拡張子チェックをスキップ
        if os.path.isdir(normalized_path):
            if not allow_directories:
                logger.warning(f"Path is a directory but directories are not allowed: {normalized_path}")
                return None
            return normalized_path
            
        # ファイルの場合は拡張子チェック
        valid_extensions = [
            '.xlsx', '.xls', '.xlsm', '.xltx', '.xltm',
            '.xml', '.mpp', '.mpt', '.pdf',
            '.html', '.htm', '.csv'
        ]
        
        file_extension = Path(normalized_path).suffix.lower()
        logger.info(f"File extension: {file_extension}")
        
        if not any(normalized_path.lower().endswith(ext) for ext in valid_extensions):
            logger.warning(f"Invalid file extension for path: {normalized_path}")
            return None
        
        # 基本的なセキュリティチェック
        if any(char in normalized_path for char in ['<', '>', '|', '"', '?', '*']):
            logger.warning(f"Invalid characters found in path: {normalized_path}")
            return None
            
        # パスの存在確認と詳細ログ
        if not os.path.exists(normalized_path):
            logger.warning(f"Path does not exist: {normalized_path}")
            # 親ディレクトリの存在確認
            parent_dir = os.path.dirname(normalized_path)
            if not os.path.exists(parent_dir):
                logger.warning(f"Parent directory does not exist: {parent_dir}")
            return None
            
        # ファイルアクセス権のチェック
        if not os.access(normalized_path, os.R_OK):
            logger.warning(f"No read permission for path: {normalized_path}")
            return None
            
        logger.info(f"Path validation successful: {normalized_path}")
        return normalized_path
        
    except Exception as e:
        logger.error(f"Error validating path {path}: {str(e)}")
        return None

def open_file_or_folder(path: str) -> Dict[str, Any]:
    """
    ファイルまたはフォルダを開く
    
    Args:
        path: 開くファイルまたはフォルダのパス
        
    Returns:
        結果を示す辞書
    """
    try:
        logger.info(f"Attempting to open path: {path}")
        is_directory = os.path.isdir(path)
        validated_path = validate_file_path(path, allow_directories=is_directory)
        
        if not validated_path:
            return {
                'success': False,
                'message': 'Invalid path specified',
                'type': 'error'
            }
        
        system = platform.system()
        result = {'success': False, 'message': '', 'type': 'error'}
        
        try:
            if system == 'Windows':
                os.startfile(validated_path)
                result = {'success': True, 'message': 'File opened successfully', 'type': 'success'}
            elif system == 'Darwin':  # macOS
                subprocess.run(['open', validated_path], check=True)
                result = {'success': True, 'message': 'File opened successfully', 'type': 'success'}
            elif system == 'Linux':
                subprocess.run(['xdg-open', validated_path], check=True)
                result = {'success': True, 'message': 'File opened successfully', 'type': 'success'}
            else:
                result = {
                    'success': False,
                    'message': f'Unsupported operating system: {system}',
                    'type': 'error'
                }
        except subprocess.CalledProcessError as e:
            logger.error(f"Process error opening file: {str(e)}")
            result = {
                'success': False,
                'message': f'Failed to open file: {str(e)}',
                'type': 'error'
            }
        except Exception as e:
            logger.error(f"Unexpected error opening file: {str(e)}")
            result = {
                'success': False,
                'message': f'Unexpected error: {str(e)}',
                'type': 'error'
            }
            
        return result
        
    except Exception as e:
        logger.error(f"Error in open_file_or_folder: {str(e)}")
        return {
            'success': False,
            'message': f'System error: {str(e)}',
            'type': 'error'
        }

def create_safe_link(path: str, text: str, allow_directories: bool = True) -> html.Button:
    """
    安全なリンクボタンの生成
    
    Args:
        path: ターゲットパス
        text: ボタンテキスト
        allow_directories: ディレクトリを許可するかどうか
        
    Returns:
        Dashボタンコンポーネント
    """
    validated_path = validate_file_path(path, allow_directories=allow_directories)
    button_id = {
        'type': 'open-path-button',
        'path': validated_path if validated_path else '',
        'action': text
    }
    
    if not validated_path:
        return html.Button(
            [
                text,
                html.Span(
                    "（ファイルが見つかりません）",
                    style={
                        'fontSize': '0.8em',
                        'color': COLORS['status']['danger'],
                        'marginLeft': '5px'
                    }
                )
            ],
            id=button_id,
            style={**STYLES['linkButton'], 'opacity': '0.5', 'cursor': 'not-allowed'},
            disabled=True
        )
    
    return html.Button(
        text,
        id=button_id,
        className='link-button',
        style=STYLES['linkButton']
    )

def create_progress_indicator(progress: float, color: str) -> html.Div:
    """プログレスバーの実装"""
    return html.Div([
        html.Div(
            style={
                **STYLES['progressBar']['bar'],
                'width': f'{progress}%',
                'backgroundColor': color,
            }
        ),
        html.Div(
            f'{progress}%',
            style=STYLES['progressBar']['text']
        )
    ], style=STYLES['progressBar']['container'])

def load_and_process_data(dashboard_file_path: str) -> pd.DataFrame:
    """
    データの読み込みと処理
    """
    try:
        # ダッシュボードデータの読み込み
        logger.info(f"Loading dashboard data from: {dashboard_file_path}")
        df = pd.read_csv(dashboard_file_path)
        
        # プロジェクトデータの読み込み
        projects_file_path = dashboard_file_path.replace('dashboard.csv', 'projects.csv')
        logger.info(f"Loading projects data from: {projects_file_path}")
        
        if not os.path.exists(projects_file_path):
            logger.error(f"Projects data file not found: {projects_file_path}")
            return df
            
        projects_df = pd.read_csv(projects_file_path)
        
        # ganttchart_pathの存在確認
        if 'ganttchart_path' not in projects_df.columns:
            logger.error("ganttchart_path column not found in projects data")
            return df
            
        # パスの検証
        projects_df['ganttchart_path'] = projects_df['ganttchart_path'].apply(
            lambda x: None if pd.isna(x) else validate_file_path(x)
        )
        
        # データの結合
        df = pd.merge(
            df,
            projects_df[['project_id', 'project_path', 'ganttchart_path']],
            on='project_id',
            how='left'
        )
        
        # 日付列の処理
        date_columns = ['task_start_date', 'task_finish_date', 'created_at']
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
            else:
                logger.warning(f"Column {col} not found in CSV")
        
        logger.info(f"Data loaded successfully. Total rows: {len(df)}")
        return df
        
    except Exception as e:
        logger.error(f"Error loading data: {str(e)}")
        return pd.DataFrame()

# データ処理関数
def check_delays(df: pd.DataFrame) -> pd.DataFrame:
    """遅延タスクの検出（改善版）"""
    current_date = datetime.datetime.now()
    return df[
        (df['task_finish_date'] < current_date) & 
        (df['task_status'] != '完了')
    ]

def get_delayed_projects_count(df: pd.DataFrame) -> int:
    """
    遅延プロジェクト数を計算（新規追加）
    遅延タスクを持つプロジェクトの数を返す
    """
    delayed_tasks = check_delays(df)
    return len(delayed_tasks['project_id'].unique())

def calculate_progress(df: pd.DataFrame) -> pd.DataFrame:
    """プロジェクト進捗の計算"""
    try:
        project_progress = df.groupby('project_id').agg({
            'project_name': 'first',
            'process': 'first',
            'line': 'first',
            'task_id': ['count', lambda x: sum(df.loc[x.index, 'task_status'] == '完了')],
            'task_milestone': lambda x: x.str.contains('○').sum(),
            'task_start_date': 'min',
            'task_finish_date': 'max',
            'project_path': 'first',
            'ganttchart_path': 'first'
        }).reset_index()
        
        project_progress.columns = [
            'project_id', 'project_name', 'process', 'line',
            'total_tasks', 'completed_tasks', 'milestone_count',
            'start_date', 'end_date', 'project_path', 'ganttchart_path'
        ]
        
        project_progress['progress'] = (project_progress['completed_tasks'] / 
                                      project_progress['total_tasks'] * 100).round(2)
        project_progress['duration'] = (project_progress['end_date'] - 
                                      project_progress['start_date']).dt.days
        
        return project_progress
    except Exception as e:
        logger.error(f"Error calculating progress: {str(e)}")
        return pd.DataFrame()

def get_status_color(progress: float, has_delay: bool) -> str:
    """進捗状況に応じた色を返す"""
    if has_delay:
        return COLORS['status']['danger']
    elif progress >= 90:
        return COLORS['status']['success']
    elif progress >= 70:
        return COLORS['status']['info']
    elif progress >= 50:
        return COLORS['status']['warning']
    return COLORS['status']['neutral']

def get_next_milestone(df: pd.DataFrame) -> pd.DataFrame:
    """次のマイルストーンを取得"""
    current_date = datetime.datetime.now()
    return df[
        (df['task_milestone'] == '○') & 
        (df['task_finish_date'] > current_date)
    ].sort_values('task_finish_date')

def next_milestone_format(next_milestones: pd.DataFrame, project_id: str) -> str:
    """マイルストーン表示のフォーマット"""
    milestone = next_milestones[next_milestones['project_id'] == project_id]
    if len(milestone) == 0:
        return '-'
    next_date = milestone.iloc[0]['task_finish_date']
    days_until = (next_date - datetime.datetime.now()).days
    return f"{milestone.iloc[0]['task_name']} ({days_until}日後)"

def get_recent_tasks(df: pd.DataFrame, project_id: str) -> html.Div:
    """
    プロジェクトの直近のタスク情報を取得し、表示用のDivを生成する
    
    Args:
        df: データフレーム
        project_id: プロジェクトID
        
    Returns:
        直近のタスク情報を含むDiv要素
    """
    try:
        current_date = datetime.datetime.now()
        project_tasks = df[df['project_id'] == project_id]
        
        # 遅延中タスク
        delayed_tasks = project_tasks[
            (project_tasks['task_finish_date'] < current_date) & 
            (project_tasks['task_status'] != '完了')
        ].sort_values('task_finish_date')
        
        # 進行中タスク（現在の日付が開始日と終了日の間にあるタスク）
        in_progress_tasks = project_tasks[
            (project_tasks['task_status'] != '完了') & 
            (project_tasks['task_start_date'] <= current_date) &
            (project_tasks['task_finish_date'] >= current_date)
        ].sort_values('task_finish_date')
        
        # 次のタスク（現在日より後に開始予定で最も近いもの）
        next_tasks = project_tasks[
            (project_tasks['task_status'] != '完了') & 
            (project_tasks['task_start_date'] > current_date)
        ].sort_values('task_start_date')
        
        # HTMLコンテンツの作成
        content_elements = []
        
        # 遅延中タスク - タスク名にも色を指定
        if len(delayed_tasks) > 0:
            content_elements.append(html.Div([
                html.Span("遅延中: ", style={'fontWeight': 'bold', 'color': COLORS['status']['danger']}),
                html.Span(delayed_tasks.iloc[0]['task_name'], style={
                    'wordBreak': 'break-word',
                    'color': COLORS['text']['primary'] # 白色を明示的に指定
                })
            ]))
        else:
            content_elements.append(html.Div([
                html.Span("遅延中: ", style={'fontWeight': 'bold', 'color': COLORS['status']['danger']}),
                html.Span("なし", style={'fontStyle': 'italic', 'color': COLORS['text']['secondary']})
            ]))
        
        # 進行中タスク - タスク名にも色を指定
        if len(in_progress_tasks) > 0:
            content_elements.append(html.Div([
                html.Span("進行中: ", style={'fontWeight': 'bold', 'color': COLORS['status']['info']}),
                html.Span(in_progress_tasks.iloc[0]['task_name'], style={
                    'wordBreak': 'break-word',
                    'color': COLORS['text']['primary'] # 白色を明示的に指定
                })
            ]))
        else:
            content_elements.append(html.Div([
                html.Span("進行中: ", style={'fontWeight': 'bold', 'color': COLORS['status']['info']}),
                html.Span("なし", style={'fontStyle': 'italic', 'color': COLORS['text']['secondary']})
            ]))
        
        # 次のタスク - タスク名にも色を指定
        if len(next_tasks) > 0:
            content_elements.append(html.Div([
                html.Span("次のタスク: ", style={'fontWeight': 'bold', 'color': COLORS['text']['accent']}),
                html.Span(next_tasks.iloc[0]['task_name'], style={
                    'wordBreak': 'break-word',
                    'color': COLORS['text']['primary'] # 白色を明示的に指定
                })
            ]))
        else:
            content_elements.append(html.Div([
                html.Span("次のタスク: ", style={'fontWeight': 'bold', 'color': COLORS['text']['accent']}),
                html.Span("なし", style={'fontStyle': 'italic', 'color': COLORS['text']['secondary']})
            ]))
        
        # 次の次のタスク - タスク名にも色を指定
        if len(next_tasks) > 1:
            content_elements.append(html.Div([
                html.Span("次の次: ", style={'fontWeight': 'bold', 'color': COLORS['text']['secondary']}),
                html.Span(next_tasks.iloc[1]['task_name'], style={
                    'wordBreak': 'break-word',
                    'color': COLORS['text']['primary'] # 白色を明示的に指定
                })
            ]))
        else:
            content_elements.append(html.Div([
                html.Span("次の次: ", style={'fontWeight': 'bold', 'color': COLORS['text']['secondary']}),
                html.Span("なし", style={'fontStyle': 'italic', 'color': COLORS['text']['secondary']})
            ]))
        
        return html.Div(content_elements, style={'fontSize': '0.9em'})
        
    except Exception as e:
        logger.error(f"Error getting recent tasks for project {project_id}: {str(e)}")
        return html.Div(
            "データ取得エラー", 
            style={'color': COLORS['status']['danger'], 'fontStyle': 'italic'}
        )

def create_project_table(df: pd.DataFrame, progress_data: pd.DataFrame) -> html.Table:
    """プロジェクト一覧テーブルの生成"""
    next_milestones = get_next_milestone(df)
    delayed_tasks = check_delays(df)
    
    rows = []
    for idx, row in progress_data.iterrows():
        has_delay = any(delayed_tasks['project_id'] == row['project_id'])
        color = get_status_color(row['progress'], has_delay)
        
        progress_indicator = create_progress_indicator(row['progress'], color)
        
        status = '遅延あり' if has_delay else '進行中' if row['progress'] < 100 else '完了'
        next_milestone = next_milestone_format(next_milestones, row['project_id'])
        task_progress = f"{row['completed_tasks']}/{row['total_tasks']}"
        
        # 直近のタスク情報を取得
        recent_tasks_content = get_recent_tasks(df, row['project_id'])
        
        # リンクボタンの生成
        links_div = html.Div([
            create_safe_link(row['project_path'], 'フォルダを開く', allow_directories=True),
            create_safe_link(row['ganttchart_path'], '工程表を開く', allow_directories=False)
        ], style={'display': 'flex', 'gap': '8px', 'justifyContent': 'center'})
        
        # ステータスセルのスタイルを設定
        status_style = {
            'padding': '10px',
            'color': COLORS['status']['danger'] if status == '遅延あり' else COLORS['text']['primary'],
            'borderBottom': '1px solid rgba(255,255,255,0.1)',
            'textAlign': 'left'
        }
        
        # 各行のセルを生成
        row_cells = [
            html.Td(row['project_name'], style={
                'padding': '10px',
                'color': COLORS['text']['primary'],
                'borderBottom': '1px solid rgba(255,255,255,0.1)',
                'textAlign': 'left',
                'minWidth': '150px'
            }),
            html.Td(row['process'], style={
                'padding': '10px',
                'color': COLORS['text']['primary'],
                'borderBottom': '1px solid rgba(255,255,255,0.1)',
                'textAlign': 'left',
                'minWidth': '100px'
            }),
            html.Td(row['line'], style={
                'padding': '10px',
                'color': COLORS['text']['primary'],
                'borderBottom': '1px solid rgba(255,255,255,0.1)',
                'textAlign': 'left',
                'minWidth': '100px'
            }),
            html.Td(progress_indicator, style={
                'padding': '10px',
                'borderBottom': '1px solid rgba(255,255,255,0.1)',
                'textAlign': 'center',
                'minWidth': '150px'
            }),
            html.Td(status, style={
                **status_style,
                'minWidth': '100px'
            }),
            html.Td(next_milestone, style={
                'padding': '10px',
                'color': COLORS['text']['primary'],
                'borderBottom': '1px solid rgba(255,255,255,0.1)',
                'textAlign': 'left',
                'minWidth': '200px'
            }),
            html.Td(task_progress, style={
                'padding': '10px',
                'color': COLORS['text']['primary'],
                'borderBottom': '1px solid rgba(255,255,255,0.1)',
                'textAlign': 'center',
                'minWidth': '100px'
            }),
            # 新しい列: 直近のタスク
            html.Td(recent_tasks_content, style={
                'padding': '10px',
                'borderBottom': '1px solid rgba(255,255,255,0.1)',
                'textAlign': 'left',
                'minWidth': '300px',
                'maxWidth': '400px'
            }),
            html.Td(links_div, style={
                'padding': '10px',
                'borderBottom': '1px solid rgba(255,255,255,0.1)',
                'textAlign': 'center',
                'minWidth': '200px',
                'whiteSpace': 'nowrap'  # ボタンが折り返されないようにする
            })
        ]
        
        rows.append(html.Tr(row_cells))
    
    return html.Table([
        html.Thead(
            html.Tr([
                html.Th(col, style={
                    'backgroundColor': COLORS['surface'],
                    'color': COLORS['text']['primary'],
                    'padding': '10px',
                    'textAlign': align,
                    'borderBottom': '1px solid rgba(255,255,255,0.1)',
                    'position': 'sticky',
                    'top': 0,
                    'zIndex': 10
                }) for col, align in zip(
                    ['プロジェクト', '工程', 'ライン', '進捗', '状態', 
                     '次のマイルストーン', 'タスク進捗', '直近のタスク', 'リンク'],
                    ['left', 'left', 'left', 'center', 'left', 
                     'left', 'center', 'left', 'center']
                )
            ])
        ),
        html.Tbody(rows)
    ], style={
        'width': '100%',
        'borderCollapse': 'collapse',
        'backgroundColor': COLORS['surface']
    })

def create_progress_distribution(progress_data: pd.DataFrame) -> go.Figure:
    """進捗状況の分布チャート作成"""
    ranges = ['0-25%', '26-50%', '51-75%', '76-99%', '100%']
    counts = [
        len(progress_data[progress_data['progress'] <= 25]),
        len(progress_data[(progress_data['progress'] > 25) & (progress_data['progress'] <= 50)]),
        len(progress_data[(progress_data['progress'] > 50) & (progress_data['progress'] <= 75)]),
        len(progress_data[(progress_data['progress'] > 75) & (progress_data['progress'] < 100)]),
        len(progress_data[progress_data['progress'] == 100])
    ]
    
    colors = COLORS['chart']['primary'][:len(ranges)]
    
    fig = go.Figure(data=[go.Bar(
        x=ranges,
        y=counts,
        marker_color=colors,
        marker_line_color='rgba(255,255,255,0.2)',
        marker_line_width=1
    )])
    
    fig.update_layout(
        **GRAPH_LAYOUT,
        margin=dict(l=40, r=20, t=20, b=40),
        height=300,
        xaxis_title='進捗率',
        yaxis_title='プロジェクト数',
        showlegend=False
    )
    
    return fig

def create_duration_distribution(progress_data: pd.DataFrame) -> go.Figure:
    """期間分布チャート作成"""
    ranges = ['1ヶ月以内', '1-3ヶ月', '3-6ヶ月', '6-12ヶ月', '12ヶ月以上']
    counts = [
        len(progress_data[progress_data['duration'] <= 30]),
        len(progress_data[(progress_data['duration'] > 30) & (progress_data['duration'] <= 90)]),
        len(progress_data[(progress_data['duration'] > 90) & (progress_data['duration'] <= 180)]),
        len(progress_data[(progress_data['duration'] > 180) & (progress_data['duration'] <= 365)]),
        len(progress_data[progress_data['duration'] > 365])
    ]
    
    fig = go.Figure(data=[go.Bar(
        x=ranges,
        y=counts,
        marker_color=COLORS['chart']['primary'][1],
        marker_line_color='rgba(255,255,255,0.2)',
        marker_line_width=1
    )])
    
    fig.update_layout(
        **GRAPH_LAYOUT,
        margin=dict(l=40, r=20, t=20, b=40),
        height=300,
        xaxis_title='プロジェクト期間',
        yaxis_title='プロジェクト数',
        showlegend=False
    )
    
    return fig

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
        
        # プロジェクト一覧（修正部分）
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

# コールバック
@app.callback(
    [Output('total-projects', 'children'),
     Output('active-projects', 'children'),
     Output('delayed-projects', 'children'),
     Output('milestone-projects', 'children'),
     Output('project-table', 'children'),
     Output('progress-distribution', 'figure'),
     Output('duration-distribution', 'figure'),
     Output('update-time', 'children')],
    [Input('update-button', 'n_clicks')]
)
def update_dashboard(n_clicks):
    try:
        # データの読み込みと処理
        df = load_and_process_data(r'C:\Users\gbrai\Documents\Projects\app_Task_Management\ProjectManager\data\exports\dashboard.csv')
        progress_data = calculate_progress(df)
        
        # 統計の計算（改善版）
        total_projects = len(progress_data)
        active_projects = len(progress_data[progress_data['progress'] < 100])
        delayed_projects = get_delayed_projects_count(df)  # 新しい関数を使用
        milestone_projects = len(df[
            (df['task_milestone'] == '○') & 
            (df['task_finish_date'].dt.month == datetime.datetime.now().month)
        ]['project_id'].unique())
        
        # テーブルとグラフの生成
        table = create_project_table(df, progress_data)
        progress_fig = create_progress_distribution(progress_data)
        duration_fig = create_duration_distribution(progress_data)
        current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        return (
            str(total_projects),
            str(active_projects),
            str(delayed_projects),
            str(milestone_projects),
            table,
            progress_fig,
            duration_fig,
            f'最終更新: {current_time}'
        )
    
    except Exception as e:
        logger.error(f"Error updating dashboard: {str(e)}")
        # エラー時のフォールバック値を返す
        return (
            '0', '0', '0', '0',
            html.Div('データの読み込みに失敗しました', style={'color': COLORS['status']['danger']}),
            go.Figure(),
            go.Figure(),
            datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        )

# ファイルを開くためのコールバック
@app.callback(
    [Output('dummy-output', 'children'),
     Output('notification-container', 'children')],
    [Input({'type': 'open-path-button', 'path': ALL, 'action': ALL}, 'n_clicks')],
    [State({'type': 'open-path-button', 'path': ALL, 'action': ALL}, 'id')]
)
def handle_button_click(n_clicks_list, button_ids):
    """
    ボタンクリックイベントの処理
    
    Args:
        n_clicks_list: クリック回数のリスト
        button_ids: ボタンIDのリスト
        
    Returns:
        通知メッセージとダミー出力
    """
    ctx = dash.callback_context
    if not ctx.triggered or not any(n_clicks_list):
        raise PreventUpdate
    
    try:
        # クリックされたボタンのインデックスを特定
        button_index = next(
            i for i, n_clicks in enumerate(n_clicks_list)
            if n_clicks is not None and n_clicks > 0
        )
        path = button_ids[button_index]['path']
        action = button_ids[button_index]['action']
        
        if not path:
            return '', html.Div(
                'Invalid path specified',
                style={**STYLES['notification']['error'], 'opacity': 1}
            )
        
        # ファイル/フォルダを開く
        result = open_file_or_folder(path)
        
        # 結果に基づいて通知を表示
        notification_style = (
            STYLES['notification']['success']
            if result['success']
            else STYLES['notification']['error']
        )
        
        return '', html.Div(
            result['message'],
            style={**notification_style, 'opacity': 1}
        )
        
    except Exception as e:
        logger.error(f"Error in callback: {str(e)}")
        return '', html.Div(
            'An error occurred',
            style={**STYLES['notification']['error'], 'opacity': 1}
        )

if __name__ == '__main__':
    app.run_server(debug=True)