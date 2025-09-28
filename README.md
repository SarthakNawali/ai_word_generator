# Enhanced AI-Powered Project Generator

A comprehensive Streamlit application that automatically generates complete academic projects with AI-powered content, professional formatting, and integrated images.

## Features

### Core Functionality
- **AI Content Generation**: Creates structured academic content using Groq's Gemma2-9B-IT model
- **Professional Word Formatting**: Generates properly formatted .docx files with headers, footers, and academic styling
- **Automatic Image Integration**: Fetches and embeds relevant images using Google Custom Search API
- **PDF Reference Processing**: Extracts content from uploaded PDFs to enhance generated content
- **Customizable Structure**: Support for both default academic sections and custom table of contents

### Document Features
- Cover page with title, author, and date
- Formal abstract generation
- Table of contents (auto-updating in Word)
- Professional formatting (Times New Roman, proper spacing, justified text)
- Headers and footers
- APA-style references
- Embedded images with captions
- Structured content with bullet points and numbered lists

## Requirements

### Python Dependencies
```
streamlit
python-docx
groq
PyPDF2
requests
Pillow (PIL)
```

### API Requirements
- **Groq API Key** (Required): For AI content generation
  - Get free API key at [Groq Console](https://console.groq.com/)
- **Google Custom Search API** (Optional): For automatic image integration
  - Google API Key from [Google Cloud Console](https://console.cloud.google.com/)
  - Custom Search Engine ID from [Google CSE](https://cse.google.com/)

## Installation

1. Clone or download the application file
2. Install required dependencies:
```bash
pip install streamlit python-docx groq PyPDF2 requests pillow
```

3. Run the application:
```bash
streamlit run app.py
```

## Usage

### Basic Setup
1. Launch the application
2. Enter your Groq API key in the sidebar
3. (Optional) Add Google API credentials for image integration

### Creating a Project
1. **Enter Project Information**:
   - Project title
   - Student name
   - Detailed project description
   - Target page count (5-50 pages)

2. **Customize Structure** (Optional):
   - Enter custom table of contents
   - Upload reference PDFs
   - Add additional notes or requirements

3. **Generate Project**:
   - Click "Generate Complete Project"
   - Wait for AI processing (typically 2-5 minutes)
   - Download the generated Word document

## Generated Document Structure

### Default Sections
- Introduction
- Literature Review
- Methodology
- Results and Analysis
- Conclusion
- References

### Document Elements
- **Cover Page**: Title, subtitle, author, date
- **Abstract**: Formal academic abstract (150-200 words)
- **Table of Contents**: Auto-updating structure
- **Main Content**: AI-generated sections with proper formatting
- **Images**: Relevant images with captions (if API enabled)
- **References**: APA-formatted citations

## Configuration Options

### Content Generation
- **Temperature**: 0.7 (balanced creativity and coherence)
- **Max Tokens**: 2000 per section
- **Model**: Gemma2-9B-IT via Groq

### Image Integration
- **Search Limit**: Up to 12 image searches per project
- **Images per Section**: Maximum 2 images
- **Image Processing**: Auto-resize, format conversion, size limits
- **Supported Formats**: JPG, PNG, JPEG

### Document Formatting
- **Font**: Times New Roman
- **Size**: 12pt body, 14pt headings
- **Spacing**: 1.15 line spacing
- **Margins**: 1 inch all sides
- **Alignment**: Justified text, centered headings

## Error Handling

The application includes comprehensive error handling for:
- Invalid API keys or quota limits
- Image download failures
- PDF processing errors
- Document formatting issues
- Network connectivity problems

## Limitations

### API Limitations
- **Groq**: Rate limits apply to free tier
- **Google Custom Search**: 100 free searches per day
- **Image Downloads**: Size and format restrictions

### Content Quality
- Generated content requires review and customization
- References are placeholder examples, not real citations
- Images are automatically selected and may need manual review

## File Structure

```
project-generator/
├── app.py                 # Main Streamlit application
├── README.md             # This documentation
└── requirements.txt      # Python dependencies (if created)
```

## Key Functions

### Core Functions
- `generate_content_with_groq()`: AI content generation
- `create_word_document_safe()`: Document creation with error handling
- `search_google_images()`: Image search and integration
- `extract_pdf_text()`: PDF content extraction

### Utility Functions
- `set_font_style()`: Font formatting
- `add_header_footer_safe()`: Header/footer creation
- `download_image_safe()`: Image processing
- `format_content_with_lists()`: Content formatting

## Troubleshooting

### Common Issues
1. **API Key Errors**: Verify keys are correct and have sufficient quota
2. **Image Integration Fails**: Check Google API setup and CSE configuration
3. **PDF Processing Errors**: Ensure PDFs contain readable text
4. **Document Formatting Issues**: May require manual adjustment in Word

### Performance Tips
- Use shorter, focused project descriptions for better results
- Limit uploaded PDFs to relevant content only
- Consider disabling image integration for faster generation

## Security Notes

- API keys are handled securely (password fields)
- Temporary files are automatically cleaned up
- No data is stored permanently by the application

## Contributing

This application is designed for educational and academic assistance purposes. When contributing:
- Maintain error handling standards
- Test with various input types
- Follow the existing code structure
- Document any new features

## Disclaimer

This tool is designed to assist with academic project creation. Users should:
- Review and customize all generated content
- Verify citations and references
- Ensure compliance with academic integrity policies
- Use generated content as a starting point, not a final submission

---

**Version**: 1.0  
**Last Updated**: September 2025  
**Compatibility**: Python 3.7+, Streamlit 1.0+
