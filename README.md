# Highlighted Text Extraction

This is a web application to extract and save highlighted text from uploaded DOCX files. The application allows users to upload DOCX files, extract highlighted text, and save it to a database. Users can view and download the extracted highlights at any time.

## Features

- Upload DOCX files and extract highlighted text.
- Save extracted highlights to a database.
- View all uploaded documents and their highlights.
- Download highlights as CSV files.
- Persistent storage of uploaded documents and highlights.

## Requirements

- Python 3.x
- Flask
- Flask-SQLAlchemy
- python-docx

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/asdman011/Highlight_Extractor.git
    cd Highlight_Extractor
    ```

2. Install the required packages:
    ```sh
    pip install flask flask_sqlalchemy python-docx
    ```

3. Run the application:
    ```sh
    python app.py
    ```

4. Open your browser and go to `http://127.0.0.1:5000`.

## Usage

1. On the homepage, upload a DOCX file to extract highlighted text.
2. View the list of uploaded documents.
3. Click "View Highlights" to see the extracted highlights of a specific document.
4. Download the highlights as a CSV file.

## Project Structure

- `app.py`: The main Flask application file.
- `models.py`: Contains the database models.
- `templates/`: Contains the HTML templates.
  - `index.html`: The homepage template.
  - `highlights.html`: The template to view highlights of a document.
- `static/`: Contains static files like CSS.
  - `styles.css`: The main stylesheet.
- `uploads/`: Directory to store uploaded DOCX files.
- `outputs/`: Directory to store output CSV files.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

## Acknowledgements

This project was inspired by the need to efficiently extract and save highlighted text from DOCX documents for various purposes.

## Contact

For any questions or feedback, please contact [your-email@example.com].

---

View the source code on [GitHub](https://github.com/asdman011/Highlight_Extractor).
