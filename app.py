import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
import plotly.graph_objects as go
import os
import time
from datetime import datetime
import tempfile
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import threading
import base64

# Page configuration
st.set_page_config(
    page_title="Excel Data Visualizer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .sidebar .sidebar-content {
        background-color: #f0f2f6;
    }
    .upload-section {
        background-color: #e8f4fd;
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 20px;
    }
    .chart-container {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

class ExcelDataHandler:
    def __init__(self):
        self.df = None
        self.file_path = None
    
    def load_data(self, uploaded_file):
        """Load data from uploaded Excel file"""
        try:
            if uploaded_file is not None:
                self.df = pd.read_excel(uploaded_file)
                return True
            return False
        except Exception as e:
            st.error(f"Error reading Excel file: {e}")
            return False
    
    def get_numeric_columns(self):
        """Get numeric columns from dataframe"""
        if self.df is not None:
            return self.df.select_dtypes(include=['number']).columns.tolist()
        return []
    
    def get_categorical_columns(self):
        """Get categorical columns from dataframe"""
        if self.df is not None:
            return self.df.select_dtypes(include=['object']).columns.tolist()
        return []
    
    def get_datetime_columns(self):
        """Get datetime columns from dataframe"""
        if self.df is not None:
            datetime_cols = []
            for col in self.df.columns:
                if pd.api.types.is_datetime64_any_dtype(self.df[col]):
                    datetime_cols.append(col)
                else:
                    # Try to convert to datetime
                    try:
                        pd.to_datetime(self.df[col])
                        datetime_cols.append(col)
                    except:
                        pass
            return datetime_cols
        return []

def create_sample_data():
    """Create sample DataFrame for demonstration"""
    sample_data = {
        'Category': ['Electronics', 'Clothing', 'Food', 'Books', 'Sports', 'Home', 'Toys'],
        'Sales': [15000, 8000, 12000, 5000, 7000, 9000, 6000],
        'Profit': [3000, 1600, 2400, 1000, 1400, 1800, 1200],
        'Month': ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul'],
        'Growth_Rate': [15, 8, 12, 5, 7, 9, 6],
        'Region': ['North', 'South', 'East', 'West', 'North', 'South', 'East']
    }
    return pd.DataFrame(sample_data)

def plot_interactive_bar(data_handler, x_col, y_col, chart_type='Bar'):
    """Create interactive bar/column chart"""
    if data_handler.df is not None:
        if chart_type == 'Bar':
            fig = px.bar(data_handler.df, x=x_col, y=y_col, 
                        title=f'{y_col} by {x_col}',
                        color=x_col,
                        template='plotly_white')
        else:  # Column chart
            fig = px.bar(data_handler.df, x=x_col, y=y_col, 
                        title=f'{y_col} by {x_col}',
                        color=x_col,
                        template='plotly_white')
        
        fig.update_layout(
            xaxis_title=x_col,
            yaxis_title=y_col,
            showlegend=False
        )
        st.plotly_chart(fig, use_container_width=True)

def plot_interactive_line(data_handler, x_col, y_col, group_col=None):
    """Create interactive line chart"""
    if data_handler.df is not None:
        if group_col and group_col in data_handler.df.columns:
            fig = px.line(data_handler.df, x=x_col, y=y_col, color=group_col,
                         title=f'{y_col} Trend by {x_col}',
                         template='plotly_white')
        else:
            fig = px.line(data_handler.df, x=x_col, y=y_col,
                         title=f'{y_col} Trend by {x_col}',
                         template='plotly_white')
        
        fig.update_layout(
            xaxis_title=x_col,
            yaxis_title=y_col
        )
        st.plotly_chart(fig, use_container_width=True)

def plot_interactive_pie(data_handler, names_col, values_col):
    """Create interactive pie chart"""
    if data_handler.df is not None:
        # Aggregate data if needed
        if data_handler.df[names_col].duplicated().any():
            pie_data = data_handler.df.groupby(names_col)[values_col].sum().reset_index()
            names_col_temp = names_col
            values_col_temp = values_col
        else:
            pie_data = data_handler.df
            names_col_temp = names_col
            values_col_temp = values_col
        
        fig = px.pie(pie_data, names=names_col_temp, values=values_col_temp,
                    title=f'Distribution of {values_col} by {names_col}',
                    template='plotly_white')
        
        fig.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig, use_container_width=True)

def plot_interactive_scatter(data_handler, x_col, y_col, color_col=None, size_col=None):
    """Create interactive scatter plot"""
    if data_handler.df is not None:
        fig = px.scatter(data_handler.df, x=x_col, y=y_col, 
                        color=color_col if color_col else None,
                        size=size_col if size_col else None,
                        title=f'{y_col} vs {x_col}',
                        template='plotly_white',
                        hover_data=data_handler.df.columns.tolist())
        
        fig.update_layout(
            xaxis_title=x_col,
            yaxis_title=y_col
        )
        st.plotly_chart(fig, use_container_width=True)

def plot_histogram(data_handler, column, bins=20):
    """Create interactive histogram"""
    if data_handler.df is not None:
        fig = px.histogram(data_handler.df, x=column, 
                          title=f'Distribution of {column}',
                          nbins=bins,
                          template='plotly_white')
        
        fig.update_layout(
            xaxis_title=column,
            yaxis_title='Frequency'
        )
        st.plotly_chart(fig, use_container_width=True)

def main():
    # Initialize session state
    if 'data_handler' not in st.session_state:
        st.session_state.data_handler = ExcelDataHandler()
    if 'data_loaded' not in st.session_state:
        st.session_state.data_loaded = False
    if 'use_sample_data' not in st.session_state:
        st.session_state.use_sample_data = False
    
    # Header
    st.markdown('<h1 class="main-header">📊 Excel Data Visualizer</h1>', unsafe_allow_html=True)
    st.markdown("Upload your Excel file and create interactive visualizations in real-time!")
    
    # Sidebar
    with st.sidebar:
        st.header("📁 Data Input")
        
        # File upload section
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
        
        # Sample data option
        use_sample = st.checkbox("Use Sample Data for Demo", value=False)
        
        if use_sample:
            st.session_state.use_sample_data = True
            st.session_state.data_handler.df = create_sample_data()
            st.session_state.data_loaded = True
            st.success("Sample data loaded successfully!")
        
        if uploaded_file is not None:
            if st.session_state.data_handler.load_data(uploaded_file):
                st.session_state.data_loaded = True
                st.session_state.use_sample_data = False
                st.success("File uploaded successfully!")
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Data info
        if st.session_state.data_loaded:
            st.header("📈 Data Information")
            st.write(f"**Rows:** {len(st.session_state.data_handler.df)}")
            st.write(f"**Columns:** {len(st.session_state.data_handler.df.columns)}")
            
            numeric_cols = st.session_state.data_handler.get_numeric_columns()
            categorical_cols = st.session_state.data_handler.get_categorical_columns()
            
            st.write(f"**Numeric Columns:** {len(numeric_cols)}")
            st.write(f"**Categorical Columns:** {len(categorical_cols)}")
    
    # Main content area
    if st.session_state.data_loaded:
        data_handler = st.session_state.data_handler
        
        # Display data preview
        st.header("🔍 Data Preview")
        with st.expander("View Raw Data"):
            st.dataframe(data_handler.df, use_container_width=True)
        
        # Statistics
        st.header("📊 Data Statistics")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.subheader("Basic Info")
            st.write(data_handler.df.info())
        
        with col2:
            st.subheader("Numeric Columns Summary")
            numeric_cols = data_handler.get_numeric_columns()
            if numeric_cols:
                st.write(data_handler.df[numeric_cols].describe())
        
        with col3:
            st.subheader("Missing Values")
            missing_data = data_handler.df.isnull().sum()
            st.write(missing_data[missing_data > 0])
        
        # Visualization section
        st.header("🎨 Create Visualizations")
        
        # Chart type selection
        chart_type = st.selectbox(
            "Select Chart Type",
            ["Bar Chart", "Line Chart", "Pie Chart", "Scatter Plot", "Histogram"]
        )
        
        # Get available columns
        numeric_cols = data_handler.get_numeric_columns()
        categorical_cols = data_handler.get_categorical_columns()
        
        # Dynamic form based on chart type
        if chart_type == "Bar Chart":
            col1, col2 = st.columns(2)
            with col1:
                x_axis = st.selectbox("X-axis (Categories)", categorical_cols)
            with col2:
                y_axis = st.selectbox("Y-axis (Values)", numeric_cols)
            
            if st.button("Generate Bar Chart"):
                with st.spinner("Creating chart..."):
                    plot_interactive_bar(data_handler, x_axis, y_axis, 'Bar')
        
        elif chart_type == "Line Chart":
            col1, col2, col3 = st.columns(3)
            with col1:
                x_axis = st.selectbox("X-axis", categorical_cols + numeric_cols)
            with col2:
                y_axis = st.selectbox("Y-axis", numeric_cols)
            with col3:
                group_by = st.selectbox("Group by (Optional)", [None] + categorical_cols)
            
            if st.button("Generate Line Chart"):
                with st.spinner("Creating chart..."):
                    plot_interactive_line(data_handler, x_axis, y_axis, group_by)
        
        elif chart_type == "Pie Chart":
            col1, col2 = st.columns(2)
            with col1:
                names_col = st.selectbox("Categories", categorical_cols)
            with col2:
                values_col = st.selectbox("Values", numeric_cols)
            
            if st.button("Generate Pie Chart"):
                with st.spinner("Creating chart..."):
                    plot_interactive_pie(data_handler, names_col, values_col)
        
        elif chart_type == "Scatter Plot":
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                x_axis = st.selectbox("X-axis", numeric_cols)
            with col2:
                y_axis = st.selectbox("Y-axis", numeric_cols)
            with col3:
                color_by = st.selectbox("Color by", [None] + categorical_cols + numeric_cols)
            with col4:
                size_by = st.selectbox("Size by", [None] + numeric_cols)
            
            if st.button("Generate Scatter Plot"):
                with st.spinner("Creating chart..."):
                    plot_interactive_scatter(data_handler, x_axis, y_axis, color_by, size_by)
        
        elif chart_type == "Histogram":
            col1, col2 = st.columns(2)
            with col1:
                column = st.selectbox("Select Column", numeric_cols)
            with col2:
                bins = st.slider("Number of Bins", 5, 100, 20)
            
            if st.button("Generate Histogram"):
                with st.spinner("Creating chart..."):
                    plot_histogram(data_handler, column, bins)
        
        # Download processed data
        st.header("💾 Export Data")
        if st.button("Download Processed Data as CSV"):
            csv = data_handler.df.to_csv(index=False)
            b64 = base64.b64encode(csv.encode()).decode()
            href = f'<a href="data:file/csv;base64,{b64}" download="processed_data.csv">Download CSV File</a>'
            st.markdown(href, unsafe_allow_html=True)
    
    else:
        # Welcome message when no data is loaded
        st.info("👆 Please upload an Excel file or use sample data to get started!")
        
        # Features overview
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("### 📈 Multiple Chart Types")
            st.write("""
            - Bar Charts
            - Line Charts  
            - Pie Charts
            - Scatter Plots
            - Histograms
            """)
        
        with col2:
            st.markdown("### 🔧 Interactive Features")
            st.write("""
            - Real-time updates
            - Interactive plots
            - Data filtering
            - Export capabilities
            """)
        
        with col3:
            st.markdown("### 📊 Data Analysis")
            st.write("""
            - Statistical summaries
            - Missing value analysis
            - Data preview
            - Column type detection
            """)

if __name__ == "__main__":
    main()