"""Streamlit UI for Excel Appender application."""

import streamlit as st
from excel_appender.core import ExcelAppender
import os

# Page configuration
st.set_page_config(
    page_title="Excel Appender",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        margin-bottom: 0.5rem;
    }
    .status-success {
        color: #28a745;
        font-weight: bold;
    }
    .status-error {
        color: #dc3545;
        font-weight: bold;
    }
    .status-info {
        color: #0066cc;
        font-weight: bold;
    }
    </style>
    """, unsafe_allow_html=True)

# Initialize session state
if 'appender' not in st.session_state:
    st.session_state.appender = ExcelAppender()
if 'status' not in st.session_state:
    st.session_state.status = None
if 'discovered_files' not in st.session_state:
    st.session_state.discovered_files = []
if 'combined_data' not in st.session_state:
    st.session_state.combined_data = None
if 'file_bytes' not in st.session_state:
    st.session_state.file_bytes = None
if 'output_filename' not in st.session_state:
    st.session_state.output_filename = None
if 'has_error' not in st.session_state:
    st.session_state.has_error = False

# Main header
st.markdown('<h1 class="main-header">📊 Excel Appender</h1>', unsafe_allow_html=True)
st.markdown("Combine multiple Excel files into one with ease!")

# Sidebar
with st.sidebar:
    st.header("📋 About")
    st.write("""
    **Excel Appender** helps you quickly merge multiple Excel files from a folder.
    
    ### How it works:
    1. Specify a folder path
    2. The app discovers all Excel files
    3. Click "Run" to combine them
    4. Download your merged file!
    """)

# Main content area
col1, col2 = st.columns([3, 1])

with col1:
    st.subheader("Step 1: Specify Folder Path")
    folder_path = st.text_input(
        "Enter folder path containing Excel files:",
        placeholder="Example: C:\\Users\\YourName\\Documents\\ExcelFiles",
        key="folder_input"
    )

# Validate and discover files
if folder_path:
    files, discovery_status = st.session_state.appender.discover_files(folder_path)
    st.session_state.discovered_files = files
    
    if files:
        st.success(f"✅ {discovery_status}")
        
        # Show file count
        st.info(f"📁 Found **{len(files)}** Excel file(s)")
        
        # Expandable file list
        with st.expander("View file list"):
            for idx, file in enumerate(files, 1):
                st.text(f"{idx}. {os.path.basename(file)}")
    else:
        st.error(f"❌ {discovery_status}")
        st.session_state.has_error = True
else:
    st.info("👉 Enter a folder path to get started")
    st.session_state.discovered_files = []

# Step 2: Run button
st.subheader("Step 2: Run Appending")

col_run, col_reset = st.columns(2)

with col_run:
    run_button = st.button(
        "🚀 Run",
        disabled=len(st.session_state.discovered_files) == 0,
        use_container_width=True,
        key="run_button"
    )

with col_reset:
    reset_button = st.button(
        "🔄 Reset",
        use_container_width=True,
        key="reset_button"
    )

# Handle reset button
if reset_button:
    st.session_state.status = None
    st.session_state.discovered_files = []
    st.session_state.combined_data = None
    st.session_state.file_bytes = None
    st.session_state.output_filename = None
    st.session_state.has_error = False
    st.rerun()

# Handle run button
if run_button and st.session_state.discovered_files:
    # Show progress
    with st.spinner("⏳ Processing Excel files..."):
        # Append files
        combined_data, append_status = st.session_state.appender.append_files(
            st.session_state.discovered_files
        )
        st.session_state.combined_data = combined_data
        st.session_state.status = append_status
        
        if combined_data is not None:
            # Generate output filename
            output_filename = st.session_state.appender.get_output_filename(
                st.session_state.discovered_files
            )
            st.session_state.output_filename = output_filename
            
            # Save to bytes
            file_bytes, save_status = st.session_state.appender.save_to_bytes(combined_data)
            st.session_state.file_bytes = file_bytes
            st.session_state.status = save_status
            st.session_state.has_error = False
        else:
            st.session_state.has_error = True

# Step 3: Show results
if st.session_state.status:
    st.subheader("Step 3: Results")
    
    if st.session_state.has_error:
        st.error(f"❌ {st.session_state.status}")
        st.info("""
        **Troubleshooting Tips:**
        - Ensure Excel files contain data
        - Check that file permissions allow reading
        - Try with a different Excel file to test
        - Make sure files are not open/locked in Excel
        """)
    else:
        st.success(f"✅ {st.session_state.status}")
        
        # Show data preview if available
        if st.session_state.combined_data is not None:
            st.subheader("📊 Data Preview")
            st.info(f"Total rows: **{len(st.session_state.combined_data)}** | Total columns: **{len(st.session_state.combined_data.columns)}**")
            
            # Display first few rows in table format (no headers)
            with st.expander("View preview (first 10 rows)"):
                preview_df = st.session_state.combined_data.head(10)
                # Create HTML table without headers
                html_table = "<table style='width:100%; border-collapse: collapse;'>"
                for _, row in preview_df.iterrows():
                    html_table += "<tr>"
                    for val in row:
                        html_table += f"<td style='border: 1px solid #ddd; padding: 8px;'>{val}</td>"
                    html_table += "</tr>"
                html_table += "</table>"
                st.markdown(html_table, unsafe_allow_html=True)
            
            # Download button
            if st.session_state.file_bytes and st.session_state.output_filename:
                st.download_button(
                    label="📥 Download Excel File",
                    data=st.session_state.file_bytes,
                    file_name=st.session_state.output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

# Footer
st.markdown("---")
st.markdown("""
    <div style='text-align: center; color: #666; font-size: 0.85rem;'>
    <p>Excel Appender v1.0 | Combine your Excel files easily</p>
    </div>
    """, unsafe_allow_html=True)
