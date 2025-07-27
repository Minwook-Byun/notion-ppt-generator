# Notion PowerPoint Generator 🎯

**Automatically generate PowerPoint presentations from Notion pages using local templates!**

A Flask web server that automatically creates presentations by leveraging local PowerPoint templates and extracting data from Notion pages.

## ✨ Features

- 🔗 **Notion Integration**: Complete auto-generation from a single Notion URL
- 📊 **Table Conversion**: Notion database → PowerPoint table automation
- 🎨 **Style Guide**: Automatic color theme and font styling
- 📋 **Template System**: Utilizes local PowerPoint templates
- 🔄 **Slide Duplication**: Template-based multi-slide generation
- 💾 **Auto Save**: Automatic presentation saving and management

## 🏗️ System Architecture

### Local Template Utilization

This project operates based on **PowerPoint template files stored locally**:

```
Template Search Paths:
📁 C:/Templates/PowerPoint/
📁 ~/Documents/PowerPoint Templates/
📁 ~/Desktop/Templates/
📁 //shared/templates/powerpoint/
```

### Template Requirements

Template files should have the following structure:

```powerpoint
Slide 1: Title Slide
├── Chapter text box (e.g., "Chapter 1")
├── Title text box (e.g., "Project Title")
└── Contents text box (e.g., "Content description")
```

**Supported File Formats:**
- `.pptx` (PowerPoint Presentation)
- `.potx` (PowerPoint Template)

## 🚀 Installation & Setup

### 1. Environment Setup

```bash
pip install -r requirements.txt
```

### 2. Template Preparation

Place PowerPoint templates locally:

```bash
# Recommended path
mkdir -p ~/Desktop/Templates
# Copy template files to the folder
cp your_template.pptx ~/Desktop/Templates/
```

### 3. Run Server

```bash
python enhanced_ppt_server.py
```

Server runs at `http://localhost:8000`

## 📋 Template Setup Guide

### Creating Template Files

1. **Create new presentation in PowerPoint**
2. **Add text boxes to the first slide:**
   ```
   📝 "Chapter" (for chapter information)
   📝 "Title" (for titles)
   📝 "Contents" (for content)
   ```
3. **Save file to template folder**

### Template Validation

```bash
# Scan templates
curl http://localhost:8000/scan-templates

# List available templates
curl http://localhost:8000/list-templates
```

## 🔗 Notion Integration Setup

### Notion Page Structure

For auto-generation, Notion pages should follow this structure:

```markdown
## 🔧 Basic Settings
**Project Name:** K-Camp Jeju 3rd Introduction
**Template:** MYSC_Sample_Template
**Font:** Pretendard
**Total Slides:** 5

## 🎨 Style Guide
**Main Color:** #1E3A8A
**Accent Color:** #F97316
**Background Color:** #F8FAFC

**Font Settings:**
- **Title:** Pretendard, 24pt, Bold
- **Body:** Pretendard, 14pt
- **Caption:** Pretendard, 12pt

## 📊 Slide Configuration
[Notion Database]
| Slide | Chapter | Title | Contents | Layout_Type |
|-------|---------|-------|----------|-------------|
| 1     | Chapter 1 | Project Intro | Content... | title_slide |
| 2     | Chapter 2 | Main Features | Content... | content_slide |

## 📋 Table Data (Optional)
[Notion Database]
| Table_ID | Parent_Slide | Row | Column | Cell_Value | Header_Type |
|----------|--------------|-----|--------|------------|-------------|
| table1   | 2           | 1   | 1      | Item       | column_header |
| table1   | 2           | 1   | 2      | Description| column_header |
```

## 🛠️ API Usage

### Basic Template Operations

```bash
# Create presentation from template
curl -X POST http://localhost:8000/create-from-template \
  -H "Content-Type: application/json" \
  -d '{
    "template_name": "MYSC_Sample_Template",
    "title": "New Presentation"
  }'

# Update slide content
curl -X POST http://localhost:8000/update-slide \
  -H "Content-Type: application/json" \
  -d '{
    "slide_number": 1,
    "chapter": "Chapter 1",
    "title": "Project Introduction",
    "contents": "Detailed content..."
  }'
```

### Notion-based Auto Generation

```bash
# Complete auto-generation from Notion URL
curl -X POST http://localhost:8000/auto-generate \
  -H "Content-Type: application/json" \
  -d '{
    "notion_url": "https://notion.so/your-page-url"
  }'
```

### File Management

```bash
# Save presentation
curl -X POST http://localhost:8000/save \
  -H "Content-Type: application/json" \
  -d '{
    "filename": "my_presentation"
  }'

# List saved files
curl http://localhost:8000/presentations
```

## ⚙️ Configuration

### Save Path Settings

Generated PPTs are saved to:

```
Windows: ~/Desktop/MyPPT/
Mac/Linux: ~/Desktop/MyPPT/
```

### Custom Template Folders

Modify paths in `enhanced_ppt_server.py`:

```python
TEMPLATE_PATHS = {
    'my_templates': Path("D:/MyTemplates"),  # Add custom path
    'common_templates': Path("C:/Templates/PowerPoint"),
    # ... other paths
}
```

## 🔧 Advanced Features

### Slide Duplication

```bash
# Duplicate first slide to create new slide
curl -X POST http://localhost:8000/duplicate-slide \
  -H "Content-Type: application/json" \
  -d '{
    "slide_number": 1,
    "new_title": "New Slide Title",
    "new_content": "New content",
    "new_chapter": "Chapter 2"
  }'
```

### Color Theme Application

```bash
# Apply color theme to presentation
curl -X POST http://localhost:8000/apply-theme \
  -H "Content-Type: application/json" \
  -d '{
    "colors": {
      "main": "#1E3A8A",
      "accent": "#F97316",
      "background": "#F8FAFC"
    }
  }'
```

## ⚠️ Important Notes

### Template Requirements

- **Local Templates Required**: This tool operates based on locally stored PowerPoint templates
- **Template Structure**: Text boxes in templates must contain keywords "Chapter", "Title", "Contents"
- **File Permissions**: Read access to template folders is required

### System Requirements

- **PowerPoint Not Required**: Uses python-pptx library, no PowerPoint installation needed
- **Memory Usage**: Be mindful of memory usage when processing large template files
- **File Size Limits**: Upload/download file size limitations apply

### Security

- **Local Environment Recommended**: Current version is for local development/testing
- **Production Deployment**: Additional security settings needed (authentication, HTTPS, file access restrictions, etc.)

## 🐛 Troubleshooting

### Template Not Found

```bash
# Verify with template scan
curl http://localhost:8000/scan-templates

# Check template file locations
ls ~/Desktop/Templates/
ls "C:/Templates/PowerPoint/"
```

### Notion Integration Errors

- Verify Notion connector is activated
- Validate Notion page structure:

```bash
curl -X POST http://localhost:8000/validate-notion \
  -H "Content-Type: application/json" \
  -d '{"notion_url": "your-notion-url"}'
```

### Memory Issues

- Check template file sizes
- Limit concurrent slide processing
- Restart server

## 🔮 Future Plans

- [ ] Support for more template layouts
- [ ] Automatic image insertion
- [ ] Auto chart/graph generation
- [ ] Template marketplace integration
- [ ] Cloud storage integration

## 📄 License

MIT License - Feel free to use, modify, and distribute.

## 🙋‍♂️ Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📞 Support

If you encounter any issues or have questions:

1. Check the troubleshooting section above
2. Open an issue on GitHub
3. Provide detailed error messages and system information

---

**⭐ If this project helps you, please give it a star!**

## 🛠️ Technical Stack

- **Backend**: Flask (Python)
- **PowerPoint Processing**: python-pptx
- **Notion Integration**: Notion API
- **Template Engine**: Custom template processing
- **File Management**: Local file system

## 📊 Example Use Cases

- **Business Presentations**: Transform Notion project docs into professional slides
- **Educational Content**: Convert lesson plans to presentation format
- **Reports**: Turn data analysis notes into visual presentations
- **Project Updates**: Automatically generate status presentations from project databases

---

*This is an experimental project. For production use, please implement additional security measures and testing.*
