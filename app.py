# !/usr/bin/python
# -*- coding: utf-8 -*-

import base64
import io

import dash
from dash import dcc
from dash import html
from dash.dependencies import Input, Output, State
import pandas as pd

import plotly.express as px
from matplotlib import pyplot as plt


import numpy as np

import openpyxl

import plotly.graph_objects as go

from dash import dash_table

import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output
from datetime import datetime as dt
from datetime import date
from datetime import timedelta

import warnings
warnings.simplefilter("ignore")

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.YETI])
server = app.server

#LAYOUT
app.layout = html.Div([
    dbc.Row([
        dbc.Col(html.Div(html.Img(src=r'\assets\images\logo.png', alt='image', style={'width': '70px'}), className='logo', style={'marginLeft': '30px'})),
        dbc.Col(html.H2('Панель мониторинга собщений', style={'color': 'black'})),
        dbc.Col(dcc.Upload(
        id='upload-data',
        children=html.Div([
            'Перетащите или выберите ',
            html.A('файл')
        ]), 
        
        className='upload',
        multiple=False), style={'display': 'flex',
                                'justifyContent': 'flex-end',
                                'marginRight': '20px'})
                ], style={'display': 'flex',
                          'alignItems': 'center',
                          'box-shadow': '0px 0px 8px rgba(0, 0, 0, 0.3)',
                          'borderBottom': '1px solid lightgrey'}, className='navbar navbar-expand-lg bg-light'),
    
    dbc.Row([
        dbc.Col(html.Div(id='card-all-msg')),
        dbc.Col(html.Div(id='card-overdue-now')),
        dbc.Col(html.Div(id='card-min')),
        dbc.Col(html.Div(id='card-mean')),
        dbc.Col(html.Div(id='card-max')),
        dbc.Col(html.Div(id='card_overdue-count'))],
               style={'display': 'flex',
                      'justify-content': 'space-between',
                      'margin-top': '30px',
                      'marginRight': '20px',
                      'marginLeft': '20px'}),
    
    dbc.Row([
        dbc.Col(html.Div(id='splot-graph')),
        dbc.Col(html.Div(id='catplot-graph'))],
               style={'margin-top': '30px'}),
    

    
    dbc.Row([
        dbc.Col(html.Div(id='stageplot-graph')),
        dbc.Col(html.Div(id='overdueplot-graph'))
    ], style={'margin-top': '30px'})
    
])



#CALLBACK
@app.callback([Output('card-all-msg', 'children'),
              Output('card-overdue-now', 'children'),
              Output('card-min', 'children'),
              Output('card-max', 'children'),
              Output('card-mean', 'children'),
              Output('card_overdue-count', 'children'),
              Output('splot-graph', 'children'),
              Output('catplot-graph', 'children'),
              Output('stageplot-graph', 'children'),
              Output('overdueplot-graph', 'children')],
              
              [Input('upload-data', 'contents')],
              [State('upload-data', 'filename')],
              prevent_initial_call=True)

#CALLBACK_FUNCTION
def update_output(contents, filename):
    if contents is not None:
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)
        
        if 'xlsx' in filename:
            df = pd.read_excel(io.BytesIO(decoded))
            df['date_diff'] = df['fact_completion_date'] - df['receipt_date']
            df['over_days_count'] = df['fact_completion_date'] - df['sheduled_completion_date']
            tbl_data = df[(df['overdue'] == 'Да') & (df['stage'] != 'Завершено')].pivot_table(index='coordinator_organization', values='overdue', aggfunc='count').reset_index().sort_values(by='overdue', ascending=False)
            tbl_data = tbl_data.rename(columns={'coordinator_organization': 'Организация', 'overdue': 'Число просрочек'})
            
            
            #CARDS
            overdue_df = df[df['over_days_count'] > '0 days']
            overdue_df = overdue_df[['coordinator_organization', 'id', 'over_days_count']].sort_values(by='coordinator_organization')
            
            card_all_msg = html.Div([html.Div('Всего сообщений', className='card-header'),
                                html.Div([
                                    html.H4(df['id'].count(), className= 'card-title')
                                ])
                            ], className='card text-white bg-success')
            overdue_now  =  df[(df['stage'] != "Завершено") & (df['overdue'] == "Да")]['id'].count()
            card_overdue_now = html.Div([
                                html.Div('Просрочено и не закрыто сейчас', className='card-header'),
                                html.Div([
                                    html.H4(overdue_now, className= 'card-title')
                                ])
                        ], className='card text-white bg-danger')
            
            min_days = str(df['date_diff'].min().days)
            card_min = html.Div([
                    html.Div('Мин. время ответа (в днях)', className='card-header'),
                        html.Div([
                            html.H4(min_days, className= 'card-title')
                                ])
                        ], className='card')
            
            mean_days = str(df['date_diff'].mean().days)
            card_mean = html.Div([
                    html.Div('Среднее время ответа (в днях)', className='card-header'),
                    html.Div([
                        html.H4(mean_days, className= 'card-title')
                            ])
                        ], className='card')
            
            max_days = str(df['date_diff'].max().days)
            card_max = html.Div([
                html.Div('Макс. время ответа (в днях)', className='card-header'),
                    html.Div([
                html.H4(max_days, className= 'card-title')
                            ])
                        ], className='card')
            
            overdue_count = df[df['overdue'] == "Да"]['id'].count()
            card_overdue_count = html.Div([
                html.Div('Всего просрочек', className='card-header'),
                    html.Div([
                html.H4(overdue_count, className= 'card-title')
                            ])
                        ], className='card')
            
            #BARS
            by_source = df.groupby(by='source')['id'].count().sort_values(ascending=True).reset_index()
            splot = px.bar(by_source, x='id', y='source', text='id')
            splot.update_layout({'plot_bgcolor': 'rgba(0, 0, 0, 0)', 'paper_bgcolor': 'rgba(0, 0, 0, 0)'}, font=dict(size=18), margin=dict(l=50, r=50, t=50, b=50), bargap=0.2, bargroupgap=0.2)
            splot.update_traces(textfont_size=25, textposition='auto')
            splot.update_xaxes(visible=True, showgrid=False, title='')
            splot.update_yaxes(visible=True, showgrid=False, title='')

            splot_graph = html.Div([html.H4('Источник сообщения', className='figure-header', style={'text-align': 'center'}),
                html.Div(dcc.Graph(figure=splot))])
            
            
            by_category = df.groupby(by='category')['id'].count().sort_values(ascending=True).reset_index().tail()
            catplot = px.bar(by_category, x='id', y='category', text='id')
            catplot.update_layout({'plot_bgcolor': 'rgba(0, 0, 0, 0)', 'paper_bgcolor': 'rgba(0, 0, 0, 0)'}, font=dict(size=18), margin=dict(l=50, r=50, t=50, b=50), bargap=0.2, bargroupgap=0.2)
            catplot.update_traces(textfont_size=25, textposition='auto')
            catplot.update_xaxes(visible=True, showgrid=False, title='')
            catplot.update_yaxes(visible=True, showgrid=False, title='')
        
            catplot_graph = html.Div([html.H4('ТОП-5 категорий', className='figure-header', style={'text-align': 'center'}),
                html.Div(dcc.Graph(figure=catplot))])
            
            
            by_stage = df.groupby(by='stage')['id'].count().sort_values(ascending=True).reset_index()

            stageplot = px.bar(by_stage,
                        y='stage',
                        x='id',
                        text='id')
            stageplot.update_layout({'plot_bgcolor': 'rgba(0, 0, 0, 0)', 'paper_bgcolor': 'rgba(0, 0, 0, 0)'}, font=dict(size=18), margin=dict(l=50, r=50, t=50, b=50), bargap=0.2, bargroupgap=0.2)
            stageplot.update_traces(textfont_size=25, textposition='auto')
            stageplot.update_xaxes(visible=True, showgrid=False, title='')
            stageplot.update_yaxes(visible=True, showgrid=False, title='')
            
            stageplot_graph = html.Div([html.H4('Этап рассмотрения', className='figure-header', style={'text-align': 'center'}),
                html.Div(dcc.Graph(figure=stageplot))])
            
            #PIE
            by_overdue = df.groupby(by='overdue')['id'].count().sort_values(ascending=False).reset_index()

            overdueplot = px.pie(by_overdue, names='overdue', values='id', hole=0.8)
            overdueplot.update_layout(legend_orientation='h', showlegend=False, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
            overdueplot.update_traces(textinfo='label+value+percent', rotation=100, textposition='auto', textfont_size=20)
            overdueplot_graph = html.Div([html.H4('Просроченные ответы', className='figure-header', style={'text-align': 'center'}),
                html.Div(dcc.Graph(figure=overdueplot))])
            
            
            
            
            
            return card_all_msg, card_overdue_now, card_min, card_mean, card_max,  card_overdue_count, splot_graph, \
                   catplot_graph, stageplot_graph, overdueplot_graph
        else:
            return html.Div('Загрузите файл в формате XLSX')
    else:
        return html.Div('Загрузите файл')

if __name__ == '__main__':
    app.run_server(debug=True, host='0.0.0.0')
    #app.run_server()
