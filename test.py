import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Cm
import io
import matplotlib.pyplot as plt
import pptx
import os
import requests
import pandas as pd
import json
from pptx.util import Pt
import sys
import matplotlib.pyplot as plt
from pptx.util import Inches
from pptx.dml.color import RGBColor
import matplotlib.colors as mcolors
from matplotlib import rcParams
import numpy as np
import seaborn as sns

st.set_page_config(page_title="Quote Responses", page_icon="💹", layout="wide")
st.html("styles.html")

st.html('<h1 class="title">Quote Responses</h1>')

# Custom CSS for styling
st.markdown("""
    <style>
    .stApp {
        background-color: #f9fcfc;
        font-family: 'Helvetica';
        font-size: 6;
    }
    .stMetric p {
        color: #007585 !important;
    }
    .stMetric {
        background-color: white;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        padding: 10px;
        border-left: 0.5rem solid #007585;
    }
    .stPlotlyChart {
        background-color: white;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        padding: 0px;
    }
    .stsubheader h2 {
        color: blue;
        font-size: 10px;
    }
    .custom-text {
        font-size: 14px;
        font-family: 'Helvetica';
    }
    .custom-text2 {
        font-size: 16px;
        font-family: 'Helvetica';
        color: #007585 !important;
    }
    .custom-text3 {
        font-size: 16px;
        font-family: 'Helvetica';
        color: #007585 !important;
    }
    .custom-text4 {
        font-size: 12px;
        font-family: 'Helvetica';
    }
    .custom-box {
        background-color: white;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        padding: 10px;
        margin-bottom: 20px;
    }
    .stButton button {
        font-family: 'Helvetica';
        font-size: 8px;
        color: #007585;
        border-radius: 5px;
        padding: 10px 10px;
    }
    .stButton button:hover {
        border: 2px solid #007585;
        font-family: 'Helvetica';
        font-size: 8px;
        color: #007585;
    }
    .stButton button:active {
        background-color: #007585;
        color: white;
        border: 2px solid #007585;
        font-family: 'Helvetica';
        font-size: 8px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="custom-text">Review the status of your quote requests and access the quotes that have been received. Received quotes can also be compared.</p>', unsafe_allow_html=True)

st.markdown('<p class="custom-text4">All displayed values are in USD ($)</p>', unsafe_allow_html=True)

data = {
    'Market': ['Ironshore', 'AIG', 'Tokio', 'Zurich', 'Swiss RE'],
    'Quote Name': ['Quote.pdf']*5,
    'Status': ["Quoted", "Quoted", "Quoted", "Quoted", "Quoted"],
    'Attachment Point': [10, 10, 20, 30, 60],
    'Limit': [10, 10, 10, 30, 30],
    '100% Layer Premium': [1, 2, 4, 5, 3],
    'Quoted Capacity': [10, 8, 8, 20, 25],
    'Signed Capacity': [5, 5, 8, 20, 25],
    'Participation Premium': [1,1,4,5,3], 
    'Bound': [True, True, True, True, False]
}

df = pd.DataFrame(data)

# Function to calculate KPIs
def calculate_kpis(df):
    total_premium = df['100% Layer Premium'].sum()
    total_capacity_value = df['Quoted Capacity'].sum()
    total_signed_premium = df['Participation Premium'].sum()
    total_signed_capacity = df['Signed Capacity'].sum()
    return total_premium, total_capacity_value, total_signed_premium, total_signed_capacity

# Editable table
gb = GridOptionsBuilder.from_dataframe(df)
gb.configure_default_column(editable=True)
grid_options = gb.build()

grid_response = AgGrid(
    df,
    gridOptions=grid_options,
    data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
    update_mode=GridUpdateMode.MODEL_CHANGED,
    fit_columns_on_grid_load=True,
    enable_enterprise_modules=True,
    height=200,
    reload_data=True,
)
st.divider()

# Update KPIs based on edited table
updated_df = pd.DataFrame(grid_response['data'])
total_premium, total_capacity_value, total_signed_premium, total_signed_capacity= calculate_kpis(updated_df)
col1, col2, col3, col4 = st.columns(4)
with col1:
    col1.metric(label="Total Premium 100%", value=f"${total_premium}", delta = "3%")

with col2:
    col2.metric(label="Total Signed Premium", value=f"${total_signed_premium}", delta = "3%")

with col3:
    col3.metric(label="Total Quoted Capacity", value=f"${total_capacity_value}", delta = "3%")

with col4:
    col4.metric(label="Total Signed Capacity", value=f"${total_signed_capacity}", delta = "3%")

# Create a new column for the combined label
updated_df['Attachment Point Label'] = updated_df['Limit'].astype(str) + " xs " + updated_df['Attachment Point'].astype(str)

# Create a capacity percentage for the plot value
updated_df['Capacity Percentage'] = (updated_df['Quoted Capacity'] / updated_df['Limit']) * 100

# Group by the new column and sum the Quoted Capacity
grouped_df = updated_df.groupby('Attachment Point Label')['Capacity Percentage'].sum().reset_index()

# Group the new column with the total signed premium
grouped_df2 = updated_df.groupby('Attachment Point Label')['Participation Premium'].sum().reset_index()

max_premium = grouped_df2['Participation Premium'].max()
min_premium = grouped_df2['Participation Premium'].min()

# Create two columns for the graphs and table
col1, col2 = st.columns(2)

with col1:
    st.markdown('<p class="custom-text2">Quoted Capacity</p>', unsafe_allow_html=True)
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=grouped_df['Attachment Point Label'],
        x=grouped_df['Capacity Percentage'],
        orientation='h',
        marker=dict(color=['red' if x <= 100 else 'green' if x==100 else 'orange' for x in grouped_df['Capacity Percentage']]),
        showlegend=False
    ))
    fig.add_shape(
        type="line",
        x0=100,
        y0=-0.5,
        x1=100,
        y1=len(grouped_df)-0.5,
        line=dict(dash="dash", width=1)
    )
    fig.update_layout(
        xaxis_title="Capacity",
        yaxis_title="Attachment Point",
        showlegend=True,
        legend=dict(
            itemsizing='constant',
            traceorder='normal',
            font=dict(size=10),
            orientation="h",
            yanchor="bottom",
            y=-0.4,
            xanchor="right",
            x=0.7
        )
    )
    fig.add_trace(go.Scatter(
        x=[None],
        y=[None],
        mode='markers',
        marker=dict(size=8, color='red'),
        name='Under Capacity'
    ))
    fig.add_trace(go.Scatter(
        x=[None],
        y=[None],
        mode='markers',
        marker=dict(size=8, color='green'),
        name='At Capacity'
    ))
    fig.add_trace(go.Scatter(
        x=[None],
        y=[None],
        mode='markers',
        marker=dict(size=8, color='orange'),
        name='Over Capacity'
    ))
    st.plotly_chart(fig, use_container_width=True)

with col2:
    st.markdown('<p class="custom-text3">Signed Premium</p>', unsafe_allow_html=True)
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=grouped_df['Attachment Point Label'],
        x=grouped_df2['Participation Premium'],
        orientation='h',
        marker=dict(color=['red' if x == max_premium else 'green' if x==min_premium else 'orange' for x in grouped_df2['Participation Premium']]),
        showlegend=False
    ))

    fig.update_layout(
        xaxis_title="Signed Premium",
        yaxis_title="Attachment Point",
        showlegend=True,
        legend=dict(
            itemsizing='constant',
            traceorder='normal',
            font=dict(size=10),
            orientation="h",
            yanchor="bottom",
            y=-0.4,
            xanchor="right",
            x=0.7
        )
    )
    fig.add_trace(go.Scatter(
        x=[None],
        y=[None],
        mode='markers',
        marker=dict(size=8, color='red'),
        name='Highest Premium'
    ))
    fig.add_trace(go.Scatter(
        x=[None],
        y=[None],
        mode='markers',
        marker=dict(size=8, color='green'),
        name='Lowest Premium'
    ))
    fig.add_trace(go.Scatter(
        x=[None],
        y=[None],
        mode='markers',
        marker=dict(size=8, color='orange'),
        name='Moderate Premium'
    ))
    st.plotly_chart(fig, use_container_width=True)

st.divider()


############################################################
############################################################
#Powerpoint Generation
############################################################
############################################################


### Summary data 

Total_placement_premium = updated_df['Participation Premium'].sum()
Theoretical_capacity = updated_df['Quoted Capacity'].sum()
Total_signed_capacity = updated_df['Signed Capacity'].sum()
Number_of_layers = updated_df['Attachment Point Label'].nunique()
Number_of_carriers = updated_df['Market'].nunique()
Submissions_sent = 30
Quotes_recieved = 16
Bound_risks = 8
Submission_to_quote_ratio = str(round((Quotes_recieved/Submissions_sent)*100, 0)) + "%"
Quote_to_bind_ratio = str(round((Bound_risks/Quotes_recieved)*100, 0)) + "%"


# Import the template presentation
presentation = pptx.Presentation("template_new.pptx")

#Function to list the text boxes and the numbers associated to the boxes for reference 
def list_text_boxes(presentation, slide_num):
    slide = presentation.slides[slide_num-1]
    text_boxes = []
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text:
            text_boxes.append(shape.text)
    return text_boxes


#Function to update the text boxes with the new values
def update_text_of_textbox(presentation, slide, text_box_id, new_text):
    slide = presentation.slides[(slide - 1)]
    count = 0
    new_text = str(new_text)
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text:
            count += 1
            if count == text_box_id:
                text_frame = shape.text_frame
                first_paragraph = text_frame.paragraphs[0]
                first_run = first_paragraph.runs[0] if first_paragraph.runs else first_paragraph.add_run()
                # Preserve formatting of the first run
                font = first_run.font
                font_name = font.name
                font_size = font.size
                font_bold = font.bold
                font_italic = font.italic
                font_underline = font.underline
                try:
                    font_color = font.color.rgb if font.color.rgb else RGBColor(209,73,0) #Default to Black
                except AttributeError:
                    font_color = RGBColor(209,73,0) #Explicit fallback to black
               
                # Clear existing text and apply new text with preserved formatting
                text_frame.clear()  # Clears all text and formatting
                new_run = text_frame.paragraphs[0].add_run()  # New run in first paragraph
                new_run.text = new_text
                # Reapply formatting
                new_run.font.name = font_name
                new_run.font.size = font_size
                new_run.font.bold = font_bold
                new_run.font.italic = font_italic
                new_run.font.underline = font_underline
                new_run.font.color.rgb = font_color # Apply fallback colour
                return



#################################################
#Executive Summary
#################################################


test_id = 'Aon has taken a strategic approach to secure the best insurance coverage for your needs. We engaged with eight insurers across four layers, ensuring a robust and competitive process. The total premium for this placement was $186,000, covering an exposure of $28 million.\n\nAs we navigated through the placement process, AXA emerged as a pivotal partner, covering 100% of the primary layer. This underscored their critical role in underwriting your most exposed risk. In the middle layers, Chubb, Hiscox, and Swiss Re shared responsibility, providing a balanced approach and reducing reliance on any single insurer. In the higher layers, AIG, Ironshore, and Apollo further diversified the risk, ensuring stability and minimizing the risk of over-reliance on one carrier.\nThe market’s perception of your risk varied across different layers. The higher layers, ranging from $20 million to $30 million, showed a significant oversupply of capacity, indicating a strong market appetite and perception of lower risk. The primary layer, from $0 to $10 million, had a balanced capacity, reflecting a stable market view.\n\nPricing dynamics played a crucial role in our strategy. The highest pricing spread was observed in the $20 million to $30 million layer at 18%, indicating diverse views among carriers and potential for negotiation. The best value was found in the $15 million to $20 million layer, offering $286 of capacity per dollar of premium, suggesting favourable market perception and cost-effectiveness.\n\nThroughout this journey, AXA, Chubb, Hiscox, and Swiss Re provided competitive pricing and capacity, making them preferred partners for future placements. AIG, Ironshore, and Apollo also contributed significantly, ensuring a well-distributed risk.\n\nIn conclusion, this placement has successfully balanced your risk across multiple carriers and layers'

update_text_of_textbox(presentation, 4, 5, test_id)


#################################################
#Summary Page 
#################################################

# Total Placement Premium 
update_text_of_textbox(presentation, 5, 16, str(Total_placement_premium))

#Theoretical Capacity
update_text_of_textbox(presentation, 5, 17, str(Theoretical_capacity))

#Total Signed Capacity 
update_text_of_textbox(presentation, 5, 18, str(Total_signed_capacity))

#Number of Layers 
update_text_of_textbox(presentation, 5, 19, str(Number_of_layers))

#Number of Carriers 
update_text_of_textbox(presentation, 5, 20, str(Number_of_carriers))

#Submissions Sent
update_text_of_textbox(presentation, 5, 21, str(Submissions_sent))

#Quotes Received
update_text_of_textbox(presentation, 5, 22, str(Quotes_recieved))

#Bound Risks
update_text_of_textbox(presentation, 5, 23, str(Bound_risks))

#Submission to Quote Ratio
update_text_of_textbox(presentation, 5, 24, str(Submission_to_quote_ratio))

#Quote to Bind Ratio
update_text_of_textbox(presentation, 5, 25, str(Quote_to_bind_ratio))


################################################
#Quote Outcome Page
################################################

plt.rcParams['figure.dpi'] = 600

# Data
data = {
    'Layer': ['Primary $10m', 'Primary $10m', 'Primary $10m', 'Primary $10m', '', 
              '\$5m excess \$10m', '\$5m excess \$10m', '\$5m excess \$10m', '\$5m excess \$10m', 
              '\$5m excess \$10m', '\$5m excess \$10m', '\$5m excess \$10m'],
    'Carriers Submitted to': ['AXA XL', 'AIG', 'Chubb', 'Beazley', '', 
                              'AXA XL', 'AIG', 'Chubb', 'AXA XL', 'AIG', 'Chubb', 'AXA XL'],
    'Quoted': ['Yes', 'No', 'Yes', 'No', '', 'Yes', 'No', 'Yes', 'Yes', 'No', 'Yes', 'No'],
    'Bound': ['Yes', 'No', 'No', 'No', '', 'Yes', 'No', 'Yes', 'No', 'No', 'No', 'No']
}

# Create DataFrame
df = pd.DataFrame(data)



# Define color map
def color_map(val):
    if val == 'Yes':
        return '#12A88A'
    elif val == 'No':
        return '#EA2238'
    else:
        return 'white'

# Create a figure and axis
fig, ax = plt.subplots(figsize=(20, 10))  # Adjust the size as needed
ax.axis('tight')
ax.axis('off')

# Create a table from the DataFrame
table = ax.table(cellText=df.values, colLabels=df.columns, cellLoc='center', loc='center')

# Apply styles to the table
for (i, j), cell in table.get_celld().items():
    

    cell.set_edgecolor('none')  # Set edge color to be the same as face color
    if i == 0:
        cell.set_text_props(weight='bold', color='white')
        cell.set_facecolor('#007c91')
        cell.set_fontsize(20)
        cell.set_height(0.05)
    elif i == 5:
        cell.set_facecolor('#d9e8ea')
        cell.set_height(0.02)
    else:
        cell.set_facecolor('white')
        cell.set_height(0.05)
        cell.set_fontsize(16)
    if j in [2, 3]:  # Apply color map to 'Quoted' and 'Bound' columns
        cell_text = cell.get_text().get_text()
        cell.set_text_props(color=color_map(cell_text))

# Save the figure as a JPEG file
plt.savefig('styled_table.png', format='jpeg', bbox_inches='tight')


def add_image_to_slide(slide, image_path, left, top, width, height):
    slide.shapes.add_picture(image_path, left, top, width, height)

add_image_to_slide(presentation.slides[5], 'styled_table.png', Cm(6.24), Cm(6.4), Cm(54.33) , Cm(27.08))



# Data for the graph
data = {
    'Company': ['AXA', 'AIG', 'Ironshore', 'Tokio', 'Apollo', 'Chubb', 'Hiscox', 'Zurich', 'Swiss Re'],
    'Signed Capacity': [33, 13, 13, 13, 7, 8, 6, 4, 3],
    'Share of Premium': [21, 26, 21, 7, 11, 5, 4, 2, 2],
    'Occurrence across Placement': [11, 12, 11, 11, 11, 11, 11, 11, 11]
}

# Create a DataFrame
df = pd.DataFrame(data)

# Set the colors for the bars
colors = ['#0072B2', '#56B4E9', '#8DD3C7']

# Set the font to Helvetica Now, font size to 16, and text color to #595959

rcParams['font.sans-serif'] = ['Helvetica']
rcParams['font.size'] = 12
rcParams['text.color'] = '#595959'
rcParams['axes.labelcolor'] = '#595959'
rcParams['xtick.color'] = '#595959'
rcParams['ytick.color'] = '#595959'
# Data for the graph
plt.rcParams['figure.dpi'] = 600
# Plot the stacked bar chart
fig, ax = plt.subplots(figsize=(10, 6))

# Plot each layer of the stack
bottom = None
for i, col in enumerate(['Signed Capacity', 'Share of Premium', 'Occurrence across Placement']):
    sns.barplot(x='Company', y=col, data=df, color=colors[i], ax=ax, label=col, bottom=bottom)
    if bottom is None:
        bottom = df[col].copy()
    else:
        bottom += df[col]

# Customize the plot to match the original
ax.set_ylabel('')
ax.set_xlabel('')
ax.legend(title='', loc='upper center', bbox_to_anchor=(0.5, 1.1), ncol=3, frameon=False)
ax.set_xticklabels(ax.get_xticklabels(), rotation=0)
ax.set_yticks([])

# Remove the data labels
for container in ax.containers:
    ax.bar_label(container, labels=['']*len(container))

# Remove borders
sns.despine(left=True, bottom=True)

# Save the figure as a png file
plt.savefig('bar.png', format='png', bbox_inches='tight')

add_image_to_slide(presentation.slides[7], 'bar.png', Cm(6.24), Cm(8.48), Cm(35.12) , Cm(23.87))



# Data for the graph
plt.rcParams['figure.dpi'] = 600

scatter_data = {
    'Company': ['AXA', 'Zurich', 'Hiscox', 'Tokio', 'Swiss Re', 'Chubb', 'AIG', 'Ironshore', 'Apollo'],
    'Premium': [80, 30, 40, 50, 20, -30, -50, -70, -90],
    'Capacity Offered': [10, 6, 4, 2, 2, -2, -4, -6, -10]
}

# Create a DataFrame
df_scatter = pd.DataFrame(scatter_data)

# Function to determine the color based on the conditions
def determine_color(premium, capacity):
    if (capacity < 0 and premium < 0) or (premium > 0 and capacity > 0):
        return '#FF9900'
    elif capacity > 0 and premium < 0:
        return '#009E73'
    else:
        return 'red'

# Apply the function to determine the color for each row
df_scatter['Color'] = df_scatter.apply(lambda row: determine_color(row['Premium'], row['Capacity Offered']), axis=1)

# Set the font to Helvetica Now, font size to 16, and text color to #595959
rcParams['font.size'] = 16
rcParams['text.color'] = '#595959'
rcParams['axes.labelcolor'] = '#595959'
rcParams['xtick.color'] = '#595959'
rcParams['ytick.color'] = '#595959'

# Create a color palette dictionary
palette = dict(zip(df_scatter['Company'], df_scatter['Color']))

# Plot the scatter plot
fig, ax = plt.subplots(figsize=(13, 9.8149606299))

# Plot each point
sns.scatterplot(x='Premium', y='Capacity Offered', hue='Company', palette=palette, data=df_scatter, s=100, ax=ax, legend=False)

# Customize the plot to match the original
ax.set_ylabel('Capacity Offered', labelpad=10)
ax.set_xlabel('Premium', labelpad=10)
ax.set_xlim(-100, 100)
ax.set_ylim(-12, 12)
rcParams['xtick.color'] = '#A6A6A6'
rcParams['ytick.color'] = '#A6A6A6'
rcParams['xtick.labelsize'] = 12
rcParams['ytick.labelsize'] = 12

# Move the spines to the center and set their colors
ax.spines['left'].set_position('zero')
ax.spines['left'].set_color('#595959')
ax.spines['left'].set_linewidth(1.5)
ax.spines['bottom'].set_position('zero')
ax.spines['bottom'].set_color('#595959')
ax.spines['bottom'].set_linewidth(1.5)

# Hide the top and right spines
ax.spines['top'].set_color('none')
ax.spines['right'].set_color('none')

# Add text labels for each point
for i in range(df_scatter.shape[0]):
    ax.text(df_scatter['Premium'].iloc[i] + 3, df_scatter['Capacity Offered'].iloc[i] - 0.1, df_scatter['Company'].iloc[i], horizontalalignment='left', size='small', color='#595959', weight='light')

# Adjust the ticks to match the original image
ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda val, pos: '{:.0f}%'.format(val)))
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda val, pos: '{:.0f}%'.format(val)))

# Set the label coordinates
ax.xaxis.set_label_coords(0.5, -0.05)
ax.yaxis.set_label_coords(-0.05, 0.5)

# Save the figure as a png file
plt.savefig('scatter.png', format='png', bbox_inches='tight')

add_image_to_slide(presentation.slides[8], 'scatter.png', Cm(32.1), Cm(8.23), Cm(31.01) , Cm(24.37))

# Function to generate PowerPoint and provide download link
def generate_pptx():
    # Save the presentation to a BytesIO object
    pptx_io = io.BytesIO()
    presentation.save(pptx_io)
    pptx_io.seek(0)
    
    # Provide download link
    st.download_button(
        label="Download PowerPoint",
        data=pptx_io,
        file_name="Client Placement Outcome.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

# Add the button to generate and download the PowerPoint
if st.button("Generate PowerPoint"):
    # Save the presentation to a BytesIO object
    pptx_io = io.BytesIO()
    presentation.save(pptx_io)
    pptx_io.seek(0)
    
    # Provide download link
    st.download_button(
        label="Download PowerPoint",
        data=pptx_io,
        file_name="Client Placement Outcome.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )