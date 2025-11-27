# AI Table Extraction System

## Project Overview
This project is an AI-based system designed to extract tables from complex, unstructured documents such as Korean application forms. It uses advanced OCR and deep learning techniques to handle nested tables, various marking styles (X, O, ✓), and unstructured layouts. The system provides a user-friendly interface for uploading documents, viewing extracted tables, and exporting results in structured formats like Excel and JSON.

## Features
- **OCR Integration**: Extracts text and tables from images, PDFs, and DOCX files.
- **Table Detection**: Handles nested tables and unstructured layouts.
- **Export Options**: Outputs results in Excel and JSON formats.
- **User Interface**: Simple and intuitive interface for document processing.

## Installation
1. Clone the repository:
   ```bash
   git clone <repository-url>
   ```
2. Navigate to the project directory:
   ```bash
   cd ai_table_extraction
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage
1. Start the application:
   ```bash
   python app.py
   ```
2. Open the application in your browser at `http://localhost:5000`.
3. Upload a document and view the extracted tables.
4. Export the results in your preferred format.

## Project Goals
- Provide a robust solution for extracting tables from complex documents.
- Ensure high accuracy and precision in table detection.
- Deliver a user-friendly experience with a modern interface.

## Challenges
- Handling nested tables and unstructured layouts.
- Supporting various marking styles (X, O, ✓).
- Ensuring compatibility with multiple document formats.

## Future Improvements
- Add support for additional languages.
- Enhance the deep learning model for better accuracy.
- Integrate cloud-based OCR services for scalability.

## License
This project is licensed under the MIT License.