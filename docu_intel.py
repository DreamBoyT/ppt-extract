import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from docx import Document
from docx.shared import Inches
import os
from PIL import Image
import io
import re

# Function to sanitize text for XML compatibility
def sanitize_text(text):
    # Adjusting the regex pattern to correctly handle Unicode ranges
    return re.sub(r'[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\u10000-\u10FFFF]', '', text)

# Function to extract flow diagrams and their shapes
def extract_flow_diagrams(shape):
    diagram_data = []
    if shape.has_text_frame and shape.text_frame.text:
        diagram_data.append(sanitize_text(shape.text_frame.text))
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for sub_shape in shape.shapes:
            diagram_data.extend(extract_flow_diagrams(sub_shape))
    return diagram_data

# Function to extract content from PPT using python-pptx
def extract_content_from_ppt(file):
    presentation = Presentation(file)
    content_with_slide_numbers = []

    for i, slide in enumerate(presentation.slides):
        slide_num = i + 1
        for shape in slide.shapes:
            # Extract images
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                image = shape.image
                image_bytes = image.blob
                content_with_slide_numbers.append((image_bytes, slide_num, 'image'))
            
            # Extract tables
            elif shape.has_table:
                table_data = []
                table = shape.table
                for row in table.rows:
                    row_data = [sanitize_text(cell.text) for cell in row.cells]
                    table_data.append(row_data)
                content_with_slide_numbers.append((table_data, slide_num, 'table'))
            
            # Extract flow diagrams and their shapes
            elif shape.shape_type in [MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.FREEFORM, MSO_SHAPE_TYPE.GROUP]:
                diagram_data = extract_flow_diagrams(shape)
                if diagram_data:
                    content_with_slide_numbers.append((diagram_data, slide_num, 'diagram'))

    return content_with_slide_numbers

# Function to create a Word document with extracted content
def create_word_document_with_content(content_with_slide_numbers, output_doc_path):
    doc = Document()
    for content, slide_num, content_type in content_with_slide_numbers:
        doc.add_paragraph(f"Slide number: {slide_num}")
        if content_type == 'image':
            img = Image.open(io.BytesIO(content))
            img_path = os.path.join("temp_image.png")
            img.save(img_path)
            doc.add_picture(img_path, width=Inches(6))  # Adjust width as needed
            os.remove(img_path)
        elif content_type == 'table':
            table = doc.add_table(rows=len(content), cols=len(content[0]))
            for row_idx, row_data in enumerate(content):
                for col_idx, cell_data in enumerate(row_data):
                    table.cell(row_idx, col_idx).text = cell_data
        elif content_type == 'diagram':
            for item in content:
                doc.add_paragraph(item)
    doc.save(output_doc_path)

# Streamlit app
st.title("PPT Image, Table, and Diagram Extractor")
uploaded_file = st.file_uploader("Choose a PPT file", type=["ppt", "pptx"])

if uploaded_file is not None:
    st.write(f"Filename: {uploaded_file.name}")
    st.write(f"File type: {uploaded_file.type}")
    st.write(f"File size: {uploaded_file.size}")

    if st.button("Extract Images, Tables, and Diagrams"):
        with st.spinner('Extracting content...'):
            content_with_slide_numbers = extract_content_from_ppt(uploaded_file)

            if content_with_slide_numbers:
                word_output_path = os.path.join(os.getcwd(), f"{os.path.splitext(uploaded_file.name)[0]}.docx")
                create_word_document_with_content(content_with_slide_numbers, word_output_path)
                st.success("Extraction successful! Click the button below to download the Word document.")
                with open(word_output_path, "rb") as word_file:
                    st.download_button(label="Download Word Document", data=word_file, file_name=f"{os.path.splitext(uploaded_file.name)[0]}.docx")
            else:
                st.error("No content found in the PPT.")
