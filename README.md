# Notion PowerPoint Generator 🎯

**Automatically generate PowerPoint presentations from Notion pages using local templates!**

An MCP (Model Context Protocol) server that automatically creates presentations by leveraging local PowerPoint templates and extracting data from Notion pages. Perfect for Claude Desktop integration and AI-powered workflows.

## ✨ Features

- 🤖 **MCP Server**: Seamless integration with Claude Desktop and AI agents
- 🔗 **Notion Integration**: Complete auto-generation from a single Notion URL
- 📊 **Table Conversion**: Notion database → PowerPoint table automation
- 🎨 **Style Guide**: Automatic color theme and font styling
- 📋 **Template System**: Utilizes local PowerPoint templates
- 🔄 **Slide Duplication**: Template-based multi-slide generation
- 💾 **Auto Save**: Automatic presentation saving and management
- 🌍 **Multilingual**: Supports both Korean and English Notion pages

## 🏗️ System Architecture

### MCP Server Design

This project is built as an **MCP (Model Context Protocol) server**, enabling:
- **Direct integration with Claude Desktop**
- **AI agent automation** - "Create a PPT from this Notion page"
- **JSON-RPC communication** for efficient data exchange
- **Tool-based architecture** with discoverable functions

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

**Key Dependencies:**
- `mcp>=1.0.0` - MCP server framework
- `python-pptx>=0.6.21` - PowerPoint processing
- `notion-client>=2.2.1` - Notion API integration

### 2. Template Preparation

Place PowerPoint templates locally:

```bash
# Recommended path
mkdir -p ~/Desktop/Templates
# Copy template files to the folder
cp your_template.pptx ~/Desktop/Templates/
```

### 3. Run MCP Server

```bash
python enhanced_ppt_server.py
```

### 4. Claude Desktop Integration

Add to your Claude Desktop MCP configuration:

```json
{
  "mcpServers": {
    "notion-ppt-generator": {
      "command": "python",
      "args": ["path/to/enhanced_ppt_server.py"],
      "env": {
        "NOTION_TOKEN": "your-notion-token"
      }
    }
  }
}
```

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

In Claude Desktop, you can now ask:
```
"Scan for available PowerPoint templates"
"List all templates in my system"
```

## 🔗 Notion Integration Setup

### Notion Page Structure

For auto-generation, Notion pages should follow this structure (supports both Korean and English):

```markdown
## 🔧 Basic Settings / 기본 설정
**Project Name / 프로젝트명:** K-Camp Jeju 3rd Introduction
**Template / 템플릿:** MYSC_Sample_Template
**Font / 폰트:** Pretendard
**Total Slides / 총 슬라이드 수:** 5

## 🎨 Style Guide / 스타일 가이드
**Main Color / 메인 컬러:** #1E3A8A
**Accent Color / 강조 컬러:** #F97316
**Background Color / 배경 컬러:** #F8FAFC

**Font Settings / 폰트 설정:**
- **Title / 제목:** Pretendard, 24pt, Bold
- **Body / 본문:** Pretendard, 14pt
- **Caption / 캡션:** Pretendard, 12pt

## 📊 Slide Configuration / 슬라이드 구성
[Notion Database]
| Slide | Chapter | Title | Contents | Layout_Type |
|-------|---------|-------|----------|-------------|
| 1     | Chapter 1 | Project Intro | Content... | title_slide |
| 2     | Chapter 2 | Main Features | Content... | content_slide |

## 📋 Table Data / 표 데이터 (Optional)
[Notion Database]
| Table_ID | Parent_Slide | Row | Column | Cell_Value | Header_Type |
|----------|--------------|-----|--------|------------|-------------|
| table1   | 2           | 1   | 1      | Item       | column_header |
| table1   | 2           | 1   | 2      | Description| column_header |
```

## 🤖 Usage with Claude Desktop

### Natural Language Commands

Once integrated with Claude Desktop, you can use natural language:

```
"Create a PowerPoint from this Notion page: [URL]"
"Generate slides using the MYSC template with chapter 'Introduction'"
"Duplicate slide 1 and update it with new content"
"Apply blue theme colors to my presentation"
"Save the current presentation as 'Q4_Report'"
```

### Available MCP Tools

- `scan_templates()` - Discover available PowerPoint templates
- `auto_generate_from_notion_url(notion_url)` - Complete automation
- `clone_template_and_update(template_name, chapter, title, contents)` - Template cloning
- `duplicate_slide(slide_number, new_title, new_content, new_chapter)` - Slide duplication
- `apply_color_theme(color_palette)` - Theme application
- `save_presentation(filename)` - File saving
- `validate_notion_structure(notion_url)` - Structure validation

## 🛠️ Advanced MCP Features

### Automated Workflows

```python
# Example: Complete automation workflow
def create_presentation_workflow(notion_url):
    # 1. Validate Notion structure
    validation = validate_notion_structure(notion_url)
    
    # 2. Auto-generate if valid
    if "PPT auto-generation possible" in validation:
        result = auto_generate_from_notion_url(notion_url)
        return result
    else:
        return "Please fix Notion page structure first"
```

### Error Handling & Fallbacks

The MCP server includes robust error handling:
- **Import errors**: Automatic fallback to minimal MCP implementation
- **Template issues**: Clear error messages with suggestions
- **Notion connectivity**: Graceful degradation when API unavailable
- **File permissions**: Helpful troubleshooting guidance

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

### MCP Server Configuration

Customize server behavior:

```python
# Server initialization with error handling
try:
    mcp = FastMCP("Enhanced PowerPoint MCP Server")
except Exception as e:
    # Fallback to minimal implementation
    mcp = MinimalMCP("Enhanced PowerPoint MCP Server")
```

## 🔧 Compatibility & Requirements

### MCP Compatibility

- **MCP Version**: 1.0.0+ supported
- **FastMCP**: Primary implementation with fallback support
- **JSON-RPC**: Standard MCP communication protocol
- **Claude Desktop**: Fully compatible

### System Requirements

- **Python**: 3.8+ recommended
- **PowerPoint**: Not required (uses python-pptx)
- **Memory**: 512MB+ available RAM
- **Storage**: 100MB+ for templates and output
- **Network**: Required for Notion API access

### Platform Support

- ✅ **Windows 10/11**
- ✅ **macOS 10.14+**
- ✅ **Linux (Ubuntu 18.04+)**
- ✅ **Docker containers**

## ⚠️ Important Notes

### MCP Server Considerations

- **Local Only**: Requires local file system access for templates
- **Claude Desktop**: Best experience when integrated with Claude Desktop
- **Session Management**: Presentations persist during MCP session
- **Concurrent Access**: Single-user design (not multi-tenant)

### Template Requirements

- **Local Templates Required**: This tool operates based on locally stored PowerPoint templates
- **Template Structure**: Text boxes in templates must contain keywords "Chapter", "Title", "Contents"
- **File Permissions**: Read access to template folders is required

### Security

- **Local Environment**: Designed for local development and personal use
- **File Access**: Requires read/write permissions for template and output directories
- **Notion API**: Store API keys securely in environment variables
- **Network**: Only connects to Notion API, no other external services

## 🐛 Troubleshooting

### MCP Server Issues

```bash
# Check MCP compatibility
python -c "from mcp.server.fastmcp import FastMCP; print('MCP OK')"

# Test server startup
python enhanced_ppt_server.py
```

### Template Discovery Problems

Ask Claude:
```
"Scan for templates and show me what's available"
"Why can't you find my PowerPoint templates?"
```

### Notion Integration Errors

- Verify Notion connector is activated in Claude Desktop
- Check Notion page structure:
```
"Validate this Notion page structure: [URL]"
```

### Memory Issues

- Check template file sizes (keep under 50MB)
- Restart Claude Desktop if experiencing slowdowns
- Clear temp files in `~/Desktop/MyPPT/`

## 🚀 Deployment Options

### Local Development

```bash
git clone https://github.com/yourusername/notion-ppt-generator
cd notion-ppt-generator
pip install -r requirements.txt
python enhanced_ppt_server.py
```

### Claude Desktop Integration

1. Clone repository
2. Add to Claude Desktop MCP configuration
3. Restart Claude Desktop
4. Test with: "What PowerPoint tools are available?"

### Docker Deployment

```dockerfile
FROM python:3.9-slim
WORKDIR /app
COPY . .
RUN pip install -r requirements.txt
CMD ["python", "enhanced_ppt_server.py"]
```

## 🔮 Future Roadmap

- [ ] **Enhanced Template Support**: More layout types and design options
- [ ] **Image Integration**: Automatic image insertion from Notion
- [ ] **Chart Generation**: Convert Notion databases to PowerPoint charts
- [ ] **Team Collaboration**: Multi-user template sharing
- [ ] **Cloud Templates**: Remote template repositories
- [ ] **Advanced Animations**: Slide transition and animation support

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

1. **GitHub Issues**: Open an issue with detailed information
2. **MCP Compatibility**: Check Claude Desktop MCP documentation
3. **Notion API**: Verify your Notion integration setup
4. **Template Problems**: Ensure templates follow required structure

---

**⭐ If this project helps you, please give it a star!**

## 🛠️ Technical Stack

- **MCP Framework**: FastMCP with fallback support
- **Backend**: Python 3.8+
- **PowerPoint Processing**: python-pptx
- **Notion Integration**: Notion API client
- **Template Engine**: Custom template processing with multilingual support
- **File Management**: Local file system with automatic organization
- **Error Handling**: Comprehensive fallback mechanisms

## 📊 Example Use Cases

- **AI-Powered Presentations**: "Claude, turn my project notes into a presentation"
- **Automated Reporting**: Convert Notion databases to executive summaries
- **Educational Content**: Transform lesson plans into engaging slides
- **Business Workflows**: Streamline pitch deck creation from project docs
- **Team Updates**: Auto-generate status presentations from sprint notes

## 🌟 Why Choose This MCP Server?

- **AI-First Design**: Built specifically for Claude Desktop integration
- **Local Control**: Your templates and data stay on your machine
- **Extensible**: Easy to customize and extend functionality
- **Robust**: Production-ready error handling and fallbacks
- **Multilingual**: Works with both Korean and English content
- **Template-Focused**: Leverages your existing PowerPoint designs

---

*This MCP server represents the cutting edge of AI-powered presentation generation. By combining Notion's organizational power with PowerPoint's presentation capabilities, it creates a seamless workflow for modern knowledge workers.*
