import streamlit as st
import os
import pandas as pd
from PIL import Image
import shutil

# Import functions from your existing invoice_processor.py
# Make sure invoice_processor.py is in the same directory as app.py
# or adjust the import path accordingly.
from invoice_processor import process_invoices, INVOICE_DIR, OUTPUT_EXCEL_FILE, ocr_image, extract_text_from_pdf, ocr_pdf_as_images, read_text_from_txt, parse_invoice_text, write_to_excel

# --- Streamlit App Configuration ---
st.set_page_config(page_title="Invoice Extractor", layout="wide")

# --- Helper Functions for Streamlit App ---
def save_uploaded_file(uploaded_file, target_path):
    """Saves an uploaded file to a target path."""
    try:
        with open(target_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return True
    except Exception as e:
        st.error(f"Error saving file {uploaded_file.name}: {e}")
        return False

def display_invoice_image(invoice_path):
    """Displays an image or the first page of a PDF."""
    if invoice_path.lower().endswith(('.png', '.jpg', '.jpeg')):
        image = Image.open(invoice_path)
        st.image(image, caption=f"Uploaded: {os.path.basename(invoice_path)}", use_column_width=True)
    elif invoice_path.lower().endswith('.pdf'):
        # For simplicity, we're not rendering PDFs directly in Streamlit here.
        # You could convert the first page to an image and display it if needed.
        st.info(f"PDF file uploaded: {os.path.basename(invoice_path)}. Preview not shown for PDFs in this basic version.")

# --- Streamlit UI ---
st.title("📄 Invoice Data Extractor")
st.markdown("""
Upload your invoice files (images, PDFs, or TXT) and this app will attempt to extract key information 
and present it in a table. You can then download the extracted data as an Excel file.
""")

# --- Global Variables / Session State ---
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'output_excel_path' not in st.session_state:
    st.session_state.output_excel_path = None
if 'temp_invoice_dir' not in st.session_state:
    # Create a unique temporary directory for this session's uploads
    # This avoids conflicts if multiple users are using a deployed app
    # For local use, a fixed subdir might be okay, but this is more robust.
    temp_dir_name = f"temp_invoices_{os.urandom(4).hex()}"
    st.session_state.temp_invoice_dir = os.path.join(".", temp_dir_name) # Create in current dir
    if not os.path.exists(st.session_state.temp_invoice_dir):
        os.makedirs(st.session_state.temp_invoice_dir, exist_ok=True)
    # Override the INVOICE_DIR from invoice_processor to use this temp one
    # This is a bit of a hack; ideally, invoice_processor would be more flexible.
    # For a cleaner approach, modify invoice_processor to accept invoice_dir as a parameter.
    globals()['INVOICE_DIR'] = st.session_state.temp_invoice_dir
    globals()['OUTPUT_EXCEL_FILE'] = os.path.join(st.session_state.temp_invoice_dir, "invoice_data_streamlit.xlsx")
if 'output_mode' not in st.session_state:
    st.session_state.output_mode = "New Excel File" # Default
if 'existing_excel_file_path' not in st.session_state:
    st.session_state.existing_excel_file_path = None
if 'existing_excel_uploader_key' not in st.session_state:
    st.session_state.existing_excel_uploader_key = 0


# --- Section for Uploading and Processing ---
with st.expander("Upload and Process Invoices", expanded=True):
    st.header("Upload Invoices")
    uploaded_files = st.file_uploader(
        "Choose invoice files",
        type=["png", "jpg", "jpeg", "pdf", "txt"],
        accept_multiple_files=True
    )

    st.header("Output Options")
    output_mode = st.radio(
        "Choose output mode:",
        ("New Excel File", "Append to Existing Excel File"),
        key='output_mode_radio',
        index=0 if st.session_state.output_mode == "New Excel File" else 1,
        on_change=lambda: st.session_state.update(output_mode=st.session_state.output_mode_radio) # Update session state on change
    )
    st.session_state.output_mode = output_mode # Ensure it's updated immediately for conditional display

    existing_excel_file_upload = None
    if st.session_state.output_mode == "Append to Existing Excel File":
        existing_excel_file_upload = st.file_uploader(
            "Upload Existing Excel File to Append to",
            type=["xlsx"],
            key=f"existing_excel_uploader_{st.session_state.existing_excel_uploader_key}" # Use key to allow re-upload
        )
        if existing_excel_file_upload:
            # Save the uploaded existing excel to a temporary location to be accessed by invoice_processor
            # This needs to be a persistent path for the processing step.
            # We'll save it in the session's temp_invoice_dir.
            existing_excel_path_in_temp = os.path.join(st.session_state.temp_invoice_dir, "existing_" + existing_excel_file_upload.name)
            if save_uploaded_file(existing_excel_file_upload, existing_excel_path_in_temp):
                st.session_state.existing_excel_file_path = existing_excel_path_in_temp
                st.success(f"Existing Excel '{existing_excel_file_upload.name}' ready for appending.")
            else:
                st.session_state.existing_excel_file_path = None # Failed to save
                st.error(f"Could not save existing Excel file: {existing_excel_file_upload.name}")


    if uploaded_files:
        st.info(f"{len(uploaded_files)} file(s) selected for processing.")
        file_details = [{"FileName": f.name, "FileType": f.type, "FileSize (bytes)": f.size} for f in uploaded_files]
        st.dataframe(file_details)

        # Save uploaded files to the session's temporary invoice directory
        saved_file_paths = []
        for uploaded_file in uploaded_files:
            target_path = os.path.join(st.session_state.temp_invoice_dir, uploaded_file.name)
            if save_uploaded_file(uploaded_file, target_path):
                saved_file_paths.append(target_path)
                st.success(f"Uploaded {uploaded_file.name}")

        if st.button("Process Uploaded Invoices", key="process"):
            if not saved_file_paths:
                st.warning("No invoice files were successfully uploaded or selected for processing.")
            elif st.session_state.output_mode == "Append to Existing Excel File" and not st.session_state.existing_excel_file_path:
                st.warning("Please upload an existing Excel file to append to, or choose 'New Excel File' mode.")
            else:
                with st.spinner("Processing invoices... This may take a moment."):
                    # Ensure the global INVOICE_DIR and OUTPUT_EXCEL_FILE are set for process_invoices
                    current_app_invoice_dir = st.session_state.temp_invoice_dir
                    # The output excel file name will be constant, but write_to_excel will handle new/append
                    current_app_output_excel = os.path.join(current_app_invoice_dir, "invoice_data_streamlit.xlsx")

                    all_extracted_data = []
                    st.write(f"Looking for invoices in: {os.path.abspath(current_app_invoice_dir)}")
                    invoice_files = os.listdir(current_app_invoice_dir)
                    # Filter out the existing excel file if it was uploaded to the same temp dir
                    invoice_files_to_process = [
                        f for f in invoice_files
                        if f != os.path.basename(current_app_output_excel) and
                           (not st.session_state.existing_excel_file_path or f != os.path.basename(st.session_state.existing_excel_file_path))
                    ]
                    st.write(f"Files found for processing: {invoice_files_to_process}")


                    for filename in invoice_files_to_process:
                        # Avoid processing the excel file if it's already there from a previous run in the same session
                        if filename == os.path.basename(current_app_output_excel):
                            continue

                        file_path = os.path.join(current_app_invoice_dir, filename)
                        extracted_text = None
                        file_type_processed = False

                        if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')):
                            st.write(f"Processing image: {filename}")
                            extracted_text = ocr_image(file_path)
                            file_type_processed = True
                        elif filename.lower().endswith('.pdf'):
                            st.write(f"Processing PDF: {filename}")
                            extracted_text = extract_text_from_pdf(file_path)
                            if not extracted_text or len(extracted_text.strip()) < 50:
                                st.write(f"Direct PDF text extraction for {filename} was minimal. Attempting OCR.")
                                extracted_text = ocr_pdf_as_images(file_path) # poppler_path might be needed
                            file_type_processed = True
                        elif filename.lower().endswith('.txt'):
                            st.write(f"Processing TXT: {filename}")
                            extracted_text = read_text_from_txt(file_path)
                            file_type_processed = True
                        
                        if file_type_processed:
                            if extracted_text and extracted_text.strip():
                                invoice_details = parse_invoice_text(extracted_text)
                                invoice_details["file_name"] = filename
                                all_extracted_data.append(invoice_details)
                            elif extracted_text is not None:
                                st.warning(f"No text could be extracted from {filename}.")
                            else:
                                st.error(f"Skipped or failed to process file: {filename}")
                        # elif filename != os.path.basename(current_app_output_excel): # Don't warn about the output excel itself
                        #     st.warning(f"Skipping unsupported file type during processing: {filename}")


                    if all_extracted_data:
                        st.session_state.processed_data = pd.DataFrame(all_extracted_data)
                        # Pass the existing excel path if in append mode
                        existing_excel_to_append = None
                        if st.session_state.output_mode == "Append to Existing Excel File":
                            existing_excel_to_append = st.session_state.existing_excel_file_path
                        
                        write_to_excel(all_extracted_data, current_app_output_excel, existing_excel_path=existing_excel_to_append)
                        st.session_state.output_excel_path = current_app_output_excel
                        st.success("Invoice processing complete!")
                    else:
                        st.session_state.processed_data = None
                        st.session_state.output_excel_path = None
                        st.warning("No data was extracted from the uploaded files.")

# --- Display Results ---
if st.session_state.processed_data is not None:
    st.subheader("Extracted Invoice Data")
    st.dataframe(st.session_state.processed_data)

    if st.session_state.output_excel_path and os.path.exists(st.session_state.output_excel_path):
        try:
            with open(st.session_state.output_excel_path, "rb") as f:
                st.download_button(
                    label="Download Data as Excel",
                    data=f,
                    file_name="invoice_data_extracted.xlsx", # User-friendly download name
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error preparing download: {e}")
    else:
        st.info("Output Excel file not found or not yet generated.")

else:
    st.info("Upload invoices and click 'Process Uploaded Invoices' to see results.")


# --- Footer & Cleanup ---
st.markdown("---")
st.markdown("Ensure Tesseract OCR and Poppler (for PDFs) are installed on the system where this app is run.")
st.markdown("Tesseract PATH in `invoice_processor.py`: `/opt/homebrew/bin/tesseract` (adjust if needed).")

# Optional: Add a button to clean up the temporary directory for the session
if st.button("Clear Session Uploads & Results"): # Moved from sidebar
    if os.path.exists(st.session_state.temp_invoice_dir):
        try:
            shutil.rmtree(st.session_state.temp_invoice_dir)
            st.success(f"Cleared temporary files in {st.session_state.temp_invoice_dir}") # Moved from sidebar
            # Reset session state related to processing
            st.session_state.processed_data = None
            st.session_state.output_excel_path = None
            st.session_state.existing_excel_file_path = None # Clear existing excel path
            st.session_state.output_mode = "New Excel File" # Reset to default
            st.session_state.existing_excel_uploader_key += 1 # Change key to reset file uploader

            # Recreate the temp dir for new uploads in the same session
            os.makedirs(st.session_state.temp_invoice_dir, exist_ok=True)
            globals()['INVOICE_DIR'] = st.session_state.temp_invoice_dir # Re-override
            globals()['OUTPUT_EXCEL_FILE'] = os.path.join(st.session_state.temp_invoice_dir, "invoice_data_streamlit.xlsx")

        except Exception as e:
            st.error(f"Error clearing temporary files: {e}") # Moved from sidebar
    st.experimental_rerun()

# To run the app:
# 1. Make sure invoice_processor.py is in the same directory as this app.py.
# 2. Open your terminal in this directory.
# 3. Run: streamlit run app.py
