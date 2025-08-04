import streamlit as st
import openai
import json
import os

# --- Configuration ---
# Make sure to set your OPENAI_API_KEY as an environment variable
# or use Streamlit's secrets management.
openai.api_key = os.getenv("OPENAI_API_KEY")

# --- Schema File Mapping ---
# Maps the name of the schema to its corresponding file.
SCHEMA_FILES = {
    "Citations": "bib-schema.txt",
    "GitHub Actions": "actions-schema.txt",
    "Resume": "resume-schema.txt"
}
SCHEMA_NAMES = list(SCHEMA_FILES.keys())
# --- Helper Functions ---

@st.cache_data
def load_schema(schema_name):
    """Loads a JSON schema from its corresponding file."""
    file_name = SCHEMA_FILES.get(schema_name)
    if not file_name:
        st.error(f"Schema file for '{schema_name}' not found in configuration.")
        return None
    
    # Assuming the script is run from the 'boom' directory's parent
    # A more robust solution would use absolute paths
    file_path = os.path.join(os.path.dirname(__file__), file_name)

    try:
        with open(file_path, 'r') as f:
            return f.read()
    except FileNotFoundError:
        st.error(f"Schema file not found at path: {file_path}")
        return None
    except Exception as e:
        st.error(f"Error loading schema file {file_name}: {e}")
        return None

def classify_text(text_input):
    """API Call 1: Classify the document type."""
    if not openai.api_key:
        st.error("OPENAI_API_KEY is not set. Please add it to your environment or Streamlit secrets.")
        st.stop()

    prompt = f"""
    You are a text classifier. Analyze the following text and identify what type of document it is.
    Respond with only one of the following words: {', '.join(SCHEMA_NAMES)}, or Unknown.

    Text to classify:
    ---
    {text_input[:1000]}
    """
    try:
        response = openai.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
            max_tokens=15  # Increased slightly to be safe
        )
        return response.choices[0].message.content.strip().replace("'", "").replace('"', '')
    except openai.OpenAIError as e:
        st.error(f"An error occurred during classification: {e}")
        return None

def extract_json(text_input, schema_name):
    """API Call 2: Extract data into the specified JSON schema."""
    schema_content = load_schema(schema_name)
    if not schema_content:
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
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": text_input}
            ],
            temperature=0.1,
        )
        json_output = response.choices[0].message.content
        # Basic validation to see if the output is likely JSON
        json.loads(json_output)
        return json_output
    except openai.OpenAIError as e:
        st.error(f"An error occurred during JSON extraction: {e}")
    except json.JSONDecodeError:
        st.error("The model returned an invalid JSON. Please try again or refine your input.")
        st.text_area("Model's Raw Output:", value=json_output, height=200)
    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")
    return None

# --- Streamlit App Layout ---
st.set_page_config(page_title="Smart JSON Converter", page_icon="✨")
st.title("✨ Smart Unstructured Text to JSON Converter")
st.write(
    "Paste your unstructured text (like a resume, BibTeX citation, or GitHub Action) "
    "and the AI will automatically detect the type and convert it to a structured JSON format."
)

input_text = st.text_area("Paste your data here:", height=350, placeholder="Example: A full resume, a BibTeX citation, or a GitHub Action workflow...")

if st.button("Convert to JSON"):
    if not input_text.strip():
        st.warning("Please paste some text to convert.")
    else:
        with st.spinner("Step 1/2: Classifying document type..."):
            doc_type = classify_text(input_text)

        if doc_type and doc_type in SCHEMA_FILES:
            st.info(f"Detected document type: **{doc_type}**")
            with st.spinner(f"Step 2/2: Converting to {doc_type} JSON..."):
                json_result = extract_json(input_text, doc_type)

            if json_result:
                st.subheader("✅ Conversion Successful!")
                st.json(json_result)
        elif doc_type == "Unknown":
            st.error("Could not determine the document type. Please try a different text or ensure it matches one of the supported types: " + ", ".join(SCHEMA_NAMES))
        else:
            st.error("Classification failed. The model returned an unexpected response. Please try again.")
            st.write("Model response:", doc_type)