# PowerPoint Arabic Converter 🔄

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/downloads/)
[![Flask](https://img.shields.io/badge/Flask-3.0.2-green.svg)](https://flask.palletsprojects.com/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A powerful web application that converts PowerPoint presentations between English and Arabic, handling RTL formatting and Arabic numerals conversion. Translation is performed client-side for better performance and reliability.

## ✨ Features

- 🔄 Client-side translation between English and Arabic
- 📝 RTL/LTR text formatting
- 🔢 Arabic numerals conversion
- 📊 Maintains presentation layout and formatting
- 🎯 Selective slide processing
- 📱 Modern, responsive UI
- ⚡ Fast processing with client-side translation

## 🚀 Quick Start

### Prerequisites

- Python 3.8 or higher
- pip (Python package manager)

### Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/powerpoint-arabic-converter.git
cd powerpoint-arabic-converter
```

2. Create and activate a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

### Running the Application

1. Start the Flask server:
```bash
python app.py
```

2. Open your browser and navigate to:
```
http://localhost:5005
```

## 💡 Usage

1. Upload your PowerPoint presentation (.pptx format)
2. Select conversion direction (English to Arabic or Arabic to English)
3. Choose specific slides to convert (optional)
4. Click "Convert" and wait for processing
5. Download the converted presentation

## 🛠️ Technical Details

### Key Components

- **Flask Backend**: Handles file uploads, processing, and downloads
- **python-pptx**: PowerPoint file manipulation
- **Client-side Translation**: Browser-based translation for better performance
- **Modern UI**: Responsive design with progress indicators

### Directory Structure

```
powerpoint-arabic-converter/
├── app.py                 # Main Flask application
├── requirements.txt       # Python dependencies
├── templates/            # HTML templates
│   └── index.html       # Main UI template
├── uploads/             # Temporary storage for uploads
└── converted/           # Output directory for converted files
```

### Features in Detail

- **Text Processing**:
  - Maintains text formatting and styles
  - Handles grouped shapes and nested elements
  - Preserves headers and footers

- **Layout Management**:
  - Mirrors slide content for RTL presentations
  - Preserves shape positioning and grouping
  - Maintains slide master elements

- **Translation**:
  - Client-side processing for better performance
  - Uses browser's translation capabilities
  - No API dependencies or rate limits

## 🔧 Configuration

The application can be configured through environment variables:

- `PORT`: Server port (default: 5005)
- `DEBUG`: Debug mode (default: False)
- `UPLOAD_FOLDER`: Upload directory path
- `CONVERTED_FOLDER`: Output directory path

## 📝 Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🤝 Support

For support, please open an issue in the GitHub repository or contact the maintainers.

## 🙏 Acknowledgments

- [python-pptx](https://python-pptx.readthedocs.io/) for PowerPoint manipulation
- [Flask](https://flask.palletsprojects.com/) for the web framework

---

Made with ❤️ for the Arabic-speaking community