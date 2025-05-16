# Invoice Data Extractor Streamlit App


This Streamlit application provides a user interface to extract data from invoice files (images, PDFs, TXT) using OCR and text parsing. The extracted data can be viewed in the app and downloaded as an Excel file. Users can choose to create a new Excel file or append data to an existing one.
Prerequisites


## Before running the application, ensure you have the following installed:
•      	Python 3.8+
•      	Tesseract OCR Engine:
–     	Follow the installation instructions for your OS: Tesseract Installation Guide
–     	Ensure Tesseract is added to your system’s PATH, or update the path in invoice_processor.py (see Configuration section).
•      	Poppler (for PDF processing with pdf2image):
–     	macOS (via Homebrew): brew install poppler
–     	Windows: Download Poppler binaries and add the bin/ directory to your PATH. Poppler for Windows
–     	Linux (Debian/Ubuntu): sudo apt-get install poppler-utils




## Setup Instructions
1.    	Clone the Repository (if applicable) or Navigate to the App Directory: bash 	# If you have the project in a git repository: 	# git clone <repository_url> 	cd "/Users/arushigupta/Desktop/Invoce to Excel/data-science-template/my-streamlit-app"
2.    	Create and Activate a Virtual Environment (Recommended): bash 	python3 -m venv venv 	# On macOS/Linux 	source venv/bin/activate 	# On Windows 	# venv\Scripts\activate
3.    	Install Dependencies: Navigate to the my-streamlit-app directory and install the required Python packages: bash 	pip install -r requirements.txt



## Running the Application
Once the setup is complete, you can run the Streamlit app:
1.    	Ensure your virtual environment is activated.
2.    	Navigate to the my-streamlit-app directory in your terminal: bash 	cd "/Users/arushigupta/Desktop/Invoce to Excel/data-science-template/my-streamlit-app"
3.    	Run the Streamlit app using the following command: bash 	streamlit run app.py This will typically open the application in your default web browser.


## How to Use the App
1.    	Upload Invoices:
–     	In the “Upload and Process Invoices” section, click “Choose invoice files” to select one or more invoice files (PNG, JPG, JPEG, PDF, TXT).
2.    	Choose Output Options:
–     	New Excel File: (Default) Extracted data will be saved to a new Excel file.
–     	Append to Existing Excel File: Select this option to add the extracted data to an existing .xlsx file. An additional file uploader will appear for you to upload the Excel file you wish to append to.
3.    	Process Invoices:
–     	Click the “Process Uploaded Invoices” button.
–     	The app will process the files, and a spinner will indicate progress.
4.    	View and Download Results:
–     	Once processing is complete, the extracted data will be displayed in a table on the main page.
–     	A “Download Data as Excel” button will appear, allowing you to download the results.
5.    	Clear Session Data:
–     	At the bottom of the page, the “Clear Session Uploads & Results” button can be used to remove all uploaded files and processed data from the current session and reset the temporary working directory.



## Configuration
•      	Tesseract OCR Path: The invoice_processor.py script requires the path to the Tesseract executable. If Tesseract is not in your system’s PATH, you need to set it manually in invoice_processor.py:
  	# Example for macOS if installed via Homebrew:
pytesseract.pytesseract.tesseract_cmd = r'/opt/homebrew/bin/tesseract'

## Example for Windows:
## pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        	Adjust this path according to your Tesseract installation location.
•      	Poppler Path (for pdf2image on Windows): If you are on Windows and pdf2image cannot find Poppler, you might need to specify the Poppler path directly in the ocr_pdf_as_images function within invoice_processor.py when calling convert_from_path: python 	# images = convert_from_path(pdf_path, poppler_path=r"C:\path\to\poppler-xx.xx.x\bin")
## Project Structure
•      	app.py: The main Streamlit application script.
•      	invoice_processor.py: Contains the core logic for OCR, text extraction, parsing, and writing to Excel.
•      	requirements.txt: Lists the Python dependencies for the Streamlit app.
temp_invoices_<session_id>/: Temporary directory created during runtime to store uploaded invoices and generated Excel files for the current session. This directory is cleaned up when the “Clear Session Uploads & Results” button is used or can be manually d
