# PowerPoint Converter

A Flask web application that converts PowerPoint presentations between English and Arabic, handling text translation, RTL formatting, and layout adjustments.

## Features

- Bidirectional translation between English and Arabic
- RTL/LTR text formatting
- Arabic numeral conversion
- Proper handling of bullet points and numbered lists
- Maintains presentation layout and formatting
- Support for selective slide conversion

## Setup

1. Clone the repository:
```bash
git clone <your-repo-url>
cd <repo-directory>
```

2. Create a virtual environment and activate it:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Run the application:
```bash
python app.py
```

The application will be available at `http://localhost:5005`

## Usage

1. Upload a PowerPoint file (.pptx)
2. Select conversion direction (English to Arabic or Arabic to English)
3. Enable/disable translation
4. Optionally specify slide numbers to convert
5. Click convert and download the processed file

## Deployment

This application is configured for deployment on Render. The deployment is automatic when pushing to the main branch.

## Project Structure

- `app.py`: Main application file
- `templates/`: HTML templates
- `uploads/`: Temporary storage for uploaded files
- `converted/`: Temporary storage for processed files

## Requirements

- Python 3.9+
- See requirements.txt for full list of dependencies

## License

MIT License #   p o w e r p o i n t _ c o n v e r t e r  
 