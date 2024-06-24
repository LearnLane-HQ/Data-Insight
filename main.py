import pandas as pd
#import snowflake.connector
import streamlit as st
from streamlit_dynamic_filters import DynamicFilters
from st_aggrid import AgGrid
from st_aggrid.grid_options_builder import GridOptionsBuilder
from st_aggrid import GridUpdateMode, DataReturnMode
import warnings
from yahooquery import Ticker

# hide future warnings (caused by st_aggrid)
warnings.simplefilter(action='ignore', category=FutureWarning)

# set page layout and define basic variables
st.set_page_config(layout="wide", page_icon='⚡', page_title="Instant Insight")
path = os.path.dirname(__file__)
today = date.today()


def resize_image(url):
    """function to resize logos while keeping aspect ratio. Accepts URL as an argument and return an image object"""

    # Open the image file
    image = Image.open(requests.get(url, stream=True).raw)

    # if a logo is too high or too wide then make the background container twice as big
    if image.height > 140:
        container_width = 220 * 2
        container_height = 140 * 2

    elif image.width > 220:
        container_width = 220 * 2
        container_height = 140 * 2
    else:
        container_width = 220
        container_height = 140

    # Create a new image with the same aspect ratio as the original image
    new_image = Image.new('RGBA', (container_width, container_height))

    # Calculate the position to paste the image so that it is centered
    x = (container_width - image.width) // 2
    y = (container_height - image.height) // 2

    # Paste the image onto the new image
    new_image.paste(image, (x, y))
    return new_image


def add_image(slide, image, left, top, width):
    """function to add an image to the PowerPoint slide and specify its position and width"""
    slide.shapes.add_picture(image, left=left, top=top, width=width)


def replace_text(replacements, slide):
    """function to replace text on a PowerPoint slide. Takes dict of {match: replacement, ... } and replaces all matches"""
    # Iterate through all shapes in the slide
    for shape in slide.shapes:
        for match, replacement in replacements.items():
            if shape.has_text_frame:
                if (shape.text.find(match)) != -1:
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        whole_text = "".join(run.text for run in paragraph.runs)
                        whole_text = whole_text.replace(str(match), str(replacement))
                        for idx, run in enumerate(paragraph.runs):
                            if idx != 0:
                                p = paragraph._p
                                p.remove(run._r)
                        if bool(paragraph.runs):
                            paragraph.runs[0].text = whole_text

def get_stock(ticker, period, interval):
    """function to get stock data from Yahoo Finance. Takes ticker, period and interval as arguments and returns a DataFrame"""
    hist = ticker.history(period=period, interval=interval)
    hist = hist.reset_index()
    # capitalize column names
    hist.columns = [x.capitalize() for x in hist.columns]
    return hist


def plot_graph(df, x, y, title, name):
    """function to plot a line graph. Takes DataFrame, x and y axis, title and name as arguments and returns a Plotly figure"""
    fig = px.line(df, x=x, y=y, template='simple_white',
                  title='<b>{} {}</b>'.format(name, title))
    fig.update_traces(line_color='#A27D4F')
    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    return fig


def peers_plot(df, name, metric):
    """function to plot a bar chart with peers. Takes DataFrame, name, metric and ticker as arguments and returns a Plotly figure"""

    # drop rows with missing metrics
    df.dropna(subset=[metric], inplace=True)

    df_sorted = df.sort_values(metric, ascending=False)

    # iterate over the labels and add the colors to the color mapping dictionary, hightlight the selected ticker
    color_map = {}
    for label in df_sorted['Company Name']:
        if label == name:
            color_map[label] = '#A27D4F'
        else:
            color_map[label] = '#D9D9D9'

    fig = px.bar(df_sorted, y='Company Name', x=metric, template='simple_white', color='Company Name',
                 color_discrete_map=color_map,
                 orientation='h',
                 title='<b>{} {} vs Peers FY22</b>'.format(name, metric))
    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', showlegend=False, yaxis_title='')
    return fig
def esg_plot(name, df):
    # Define colors for types
    colors = {name: '#A27D4F', 'Peer Group': '#D9D9D9'}

    # Creating the bar chart
    fig = go.Figure()
    for type in df['Type'].unique():
        fig.add_trace(go.Bar(
            x=df[df['Type'] == type]['variable'],
            y=df[df['Type'] == type]['value'],
            name=type,
            text=df[df['Type'] == type]['value'],
            textposition='outside',
            marker_color=colors[type]
        ))
    fig.update_layout(
        height=700,
        width=1000,
        barmode='group',
        title="ESG Score vs Peers Average",
        xaxis_title="",
        yaxis_title="Score",
        legend_title="Type",
        xaxis=dict(tickangle=0),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)')
    return fig


def get_financials(df, col_name, metric_name):
    """function to get financial metrics from a DataFrame. Takes DataFrame, column name and metric name as arguments and returns a DataFrame"""
    metric = df.loc[:, ['asOfDate', col_name]]
    metric_df = pd.DataFrame(metric).reset_index()
    metric_df.columns = ['Symbol', 'Year', metric_name]

    return metric_df
