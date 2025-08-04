import streamlit as st
import openai
import json
import os
from PyPDF2 import PdfReader
from io import StringIO

# --- Configuration ---
# Make sure to set your OPENAI_API_KEY as an environment variable
# or use Streamlit's secrets management.
openai.api_key = os.getenv("OPENAI_API_KEY")

# --- Helper Functions ---

def extract_text_from_pdf(file_obj):
    """Extracts text from an uploaded PDF file."""
    try:
        pdf_reader = PdfReader(file_obj)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text()
        return text
    except Exception as e:
        st.error(f"Error reading PDF file: {e}")
        return None

def convert_to_json(text_input, schema_content):
    """Calls OpenAI API to convert text to the specified JSON schema."""
    if not openai.api_key:
        st.error("OPENAI_API_KEY is not set. Please add it to your environment or Streamlit secrets.")
        st.stop()

    # Validate that the schema is valid JSON
    try:
        json.loads(schema_content)
    except json.JSONDecodeError:
        st.error("The provided schema is not valid JSON. Please check the format.")
        st.text_area("Invalid Schema:", value=schema_content, height=200)
        return None

    system_prompt = f"""
You are an expert data extraction tool. Your task is to convert the user's text into a valid JSON object
that strictly adheres to the following JSON schema. Do not add any fields that are not in the schema.
Ensure all required fields are present. Output ONLY the JSON object, with no additional text or explanations.

JSON Schema:
{schema_content}
    """
    try:
        response = openai.chat.completions.create(
            model="gpt-4-1106-preview",
            messages=[{"role": "system", "content": system_prompt},
                      {"role": "user", "content": text_input}],
            temperature=0.1,
        )
        json_output = response.choices[0].message.content
        # Final validation to see if the output is likely JSON
        json.loads(json_output)
        return json_output
    except openai.OpenAIError as e:
        st.error(f"An error occurred with the OpenAI API: {e}")
    except json.JSONDecodeError:
        st.error("The model returned an invalid JSON. Please try again or refine your input.")
        st.text_area("Model's Raw Output:", value=json_output, height=200)
    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")
    return None

# --- Streamlit App Layout ---
st.set_page_config(page_title="Universal JSON Converter", page_icon="⚙️")
st.title("⚙️ Universal Text-to-JSON Converter")
st.write(
    "Upload your data (text or PDF) and a JSON schema, and the AI will perform the conversion. "
    "File uploads take priority over pasted text."
)

# --- UI Columns ---
col1, col2 = st.columns(2)

# --- Column 1: Input Data ---
with col1:
    st.subheader("1. Provide Input Data")
    input_file = st.file_uploader(
        "Upload a data file (.txt or .pdf)",
        type=['txt', 'pdf']
    )
    input_text_area = st.text_area(
        "Or paste raw text here:",
        height=300,
        placeholder="John Doe\nSoftware Engineer at Acme Inc..."
    )

# --- Column 2: Target Schema ---
with col2:
    st.subheader("2. Provide JSON Schema")
    schema_file = st.file_uploader(
        "Upload a schema file (.json or .txt)",
        type=['json', 'txt']
    )
    schema_text_area = st.text_area(
        "Or paste schema here:",
        height=300,
        placeholder='{\n  "name": "string",\n  "company": "string"\n}'
    )

# --- Conversion Button ---
if st.button("Convert to JSON", type="primary"):
    # Determine final input text
    final_input_text = ""
    if input_file:
        if input_file.type == "application/pdf":
            final_input_text = extract_text_from_pdf(input_file)
        else: # txt file
            stringio = StringIO(input_file.getvalue().decode("utf-8"))
            final_input_text = stringio.read()
    elif input_text_area.strip():
        final_input_text = input_text_area
    else:
        st.warning("Please upload or paste your input data.")

    # Determine final schema content
    final_schema_content = ""
    if schema_file:
        stringio = StringIO(schema_file.getvalue().decode("utf-8"))
        final_schema_content = stringio.read()
    elif schema_text_area.strip():
        final_schema_content = schema_text_area
    else:
        st.warning("Please upload or paste a target JSON schema.")

    # --- Run Conversion ---
    if final_input_text and final_schema_content:
        with st.spinner("Converting... Please wait."):
            json_result = convert_to_json(final_input_text, final_schema_content)
            if json_result:
                st.subheader("✅ Conversion Successful!")
                st.json(json_result)
    else:
        # Warnings are already displayed, just preventing the spinner
        pass
