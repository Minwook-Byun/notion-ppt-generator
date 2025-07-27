# Enhanced MCP PowerPoint Setup - Template Clone and Update Version

Write-Host "=== Enhanced MCP PowerPoint Setup (Template Clone and Update) ===" -ForegroundColor Green

# Step 1: Clean up existing setup
Write-Host "`n1. Cleaning up existing setup..." -ForegroundColor Cyan
$claudeDir = "$env:APPDATA\Claude"
if (Test-Path "$claudeDir\claude_desktop_config.json") {
    Remove-Item "$claudeDir\claude_desktop_config.json" -Force
    Write-Host "Removed existing config file" -ForegroundColor Green
}

# Clean up existing installations
Remove-Item "C:\MCP-PowerPoint*" -Recurse -Force -ErrorAction SilentlyContinue

# Step 2: Create working directory
Write-Host "`n2. Setting up working directory..." -ForegroundColor Cyan
$workDir = "C:\MCP-PowerPoint-Enhanced"
New-Item -ItemType Directory -Path $workDir -Force | Out-Null
Set-Location $workDir

# Step 3: Setup Python virtual environment and packages
Write-Host "`n3. Setting up Python environment..." -ForegroundColor Cyan
python -m venv venv

# Activate virtual environment and install packages
Write-Host "Installing packages..." -ForegroundColor Yellow
& "$workDir\venv\Scripts\python.exe" -m pip install --upgrade pip
& "$workDir\venv\Scripts\python.exe" -m pip install python-pptx
& "$workDir\venv\Scripts\python.exe" -m pip install mcp
& "$workDir\venv\Scripts\python.exe" -m pip install fastmcp
& "$workDir\venv\Scripts\python.exe" -m pip install requests
Write-Host "Python packages installed successfully" -ForegroundColor Green

# Step 4: Configure Claude Desktop
Write-Host "`n4. Configuring Claude Desktop..." -ForegroundColor Cyan
if (!(Test-Path $claudeDir)) {
    New-Item -ItemType Directory -Path $claudeDir -Force | Out-Null
}

$pythonPath = "$workDir\venv\Scripts\python.exe" -replace '\\', '/'
$serverPath = "$workDir\enhanced_ppt_server.py" -replace '\\', '/'

# Create config object
$config = @{
    mcpServers = @{
        "enhanced-powerpoint" = @{
            command = $pythonPath
            args = @($serverPath)
            env = @{}
        }
    }
}

# Convert to JSON and save
$configJson = $config | ConvertTo-Json -Depth 4
[System.IO.File]::WriteAllText("$claudeDir\claude_desktop_config.json", $configJson, [System.Text.UTF8Encoding]::new($false))
Write-Host "Claude Desktop configuration completed" -ForegroundColor Green

# Step 5: Create template directories
Write-Host "`n5. Setting up template directories..." -ForegroundColor Cyan

$templateDirs = @(
    "C:\Templates\PowerPoint",
    "$env:USERPROFILE\Documents\PowerPoint Templates",
    "$env:USERPROFILE\Desktop\Templates"
)

foreach ($dir in $templateDirs) {
    try {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Host "Created template directory: $dir" -ForegroundColor Green
    } catch {
        Write-Host "Could not create: $dir" -ForegroundColor Yellow
    }
}

# Step 6: Create sample template with placeholder text
Write-Host "`n6. Creating sample template with placeholders..." -ForegroundColor Cyan
$sampleTemplateScript = @'
import os
from pptx import Presentation
from pptx.util import Pt, Inches
from pathlib import Path

def create_sample_template():
    """Create a sample PowerPoint template with proper placeholders"""
    try:
        prs = Presentation()
        
        # Use a title and content layout
        title_slide_layout = prs.slide_layouts[0]  # Title slide
        slide = prs.slides.add_slide(title_slide_layout)
        
        # Set title placeholder
        if slide.shapes.title:
            slide.shapes.title.text = "Title"
        
        # Add subtitle if available
        if len(slide.placeholders) > 1:
            slide.placeholders[1].text = "Chapter"
        
        # Add content slide
        content_layout = prs.slide_layouts[1]  # Title and content
        content_slide = prs.slides.add_slide(content_layout)
        
        # Set content slide title
        if content_slide.shapes.title:
            content_slide.shapes.title.text = "Title"
        
        # Set content placeholder
        if len(content_slide.placeholders) > 1:
            content_slide.placeholders[1].text = "Contents"
        
        # Save to templates directory
        template_path = Path.home() / "Desktop" / "Templates" / "YOUR COMPANY_Sample_Template.pptx"
        template_path.parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(template_path))
        
        print(f"Sample template created: {template_path}")
        
        # Also create a more complex template
        prs2 = Presentation()
        
        # Title slide
        title_slide = prs2.slides.add_slide(prs2.slide_layouts[0])
        if title_slide.shapes.title:
            title_slide.shapes.title.text = "Title"
        if len(title_slide.placeholders) > 1:
            title_slide.placeholders[1].text = "Chapter"
        
        # Content slide with manual text boxes
        blank_layout = prs2.slide_layouts[6]  # Blank slide
        manual_slide = prs2.slides.add_slide(blank_layout)
        
        # Add manual text boxes with placeholder text
        chapter_box = manual_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(1))
        chapter_box.text_frame.text = "Chapter"
        
        title_box = manual_slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(8), Inches(1))
        title_box.text_frame.text = "Title"
        
        content_box = manual_slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(8), Inches(4))
        content_box.text_frame.text = "Contents"
        
        # Save advanced template
        advanced_template_path = Path.home() / "Desktop" / "Templates" / "Advanced_Template.pptx"
        prs2.save(str(advanced_template_path))
        
        print(f"Advanced template created: {advanced_template_path}")
        return True
    except Exception as e:
        print(f"Error creating sample template: {e}")
        return False

if __name__ == "__main__":
    create_sample_template()
'@

$sampleTemplateScript | Out-File -FilePath "$workDir\create_sample_template.py" -Encoding UTF8
try {
    & "$workDir\venv\Scripts\python.exe" "$workDir\create_sample_template.py"
    Write-Host "Sample templates created successfully" -ForegroundColor Green
} catch {
    Write-Host "Could not create sample templates" -ForegroundColor Yellow
}

# Step 7: Validation
Write-Host "`n7. Validating installation..." -ForegroundColor Cyan

# JSON validation
try {
    $testConfig = Get-Content "$claudeDir\claude_desktop_config.json" -Raw | ConvertFrom-Json
    Write-Host "JSON configuration file is valid" -ForegroundColor Green
} catch {
    Write-Host "JSON configuration error: $($_.Exception.Message)" -ForegroundColor Red
}

# Test Python environment
try {
    $testResult = & "$workDir\venv\Scripts\python.exe" -c "import pptx; import mcp; import fastmcp; import requests; print('All packages working')"
    Write-Host $testResult -ForegroundColor Green
} catch {
    Write-Host "Python environment error" -ForegroundColor Red
}

Write-Host "`nInstallation completed!" -ForegroundColor Green

# Final instructions
Write-Host "`nEnhanced PowerPoint MCP Setup Complete!" -ForegroundColor Green
Write-Host "`nInstalled locations:" -ForegroundColor White
Write-Host "Server file: $workDir\enhanced_ppt_server.py" -ForegroundColor Gray
Write-Host "Python environment: $workDir\venv\" -ForegroundColor Gray
Write-Host "Config file: $claudeDir\claude_desktop_config.json" -ForegroundColor Gray

Write-Host "`nTemplate directories created:" -ForegroundColor White
foreach ($dir in $templateDirs) {
    Write-Host "$dir" -ForegroundColor Gray
}

Write-Host "`nSample templates created:" -ForegroundColor White
Write-Host "Sample_Template.pptx" -ForegroundColor Gray
Write-Host "Advanced_Template.pptx" -ForegroundColor Gray

Write-Host "`nNext steps:" -ForegroundColor White
Write-Host "1. Copy the Python server code to: $workDir\enhanced_ppt_server.py" -ForegroundColor Yellow
Write-Host "2. Place your template files in any template directory" -ForegroundColor Yellow
Write-Host "3. Completely close Claude Desktop (check Task Manager)" -ForegroundColor Yellow
Write-Host "4. Start Claude Desktop again" -ForegroundColor Yellow
Write-Host "5. Start a new conversation" -ForegroundColor Yellow
Write-Host "6. Test: scan_templates" -ForegroundColor Yellow
Write-Host "7. Test: clone_template_and_update Sample_Template Chapter1 AI_Future AI_Technology_Trends" -ForegroundColor Yellow

Write-Host "`nCore Features:" -ForegroundColor White
Write-Host "clone_template_and_update - Clone template and update text (CORE FEATURE)" -ForegroundColor Cyan
Write-Host "update_specific_slide_text - Update specific slide text" -ForegroundColor Cyan
Write-Host "scan_templates - Discover available templates" -ForegroundColor Cyan
Write-Host "list_available_templates - Show all template list" -ForegroundColor Cyan
Write-Host "create_presentation_from_template - Create presentation from template" -ForegroundColor Cyan
Write-Host "save_presentation - Save presentation" -ForegroundColor Cyan
Write-Host "get_presentation_info - Current presentation info" -ForegroundColor Cyan

Write-Host "`nUsage Examples:" -ForegroundColor White
Write-Host "scan_templates" -ForegroundColor Gray
Write-Host "clone_template_and_update Sample_Template Chapter1 AI_Technology AI_trends_and_analysis" -ForegroundColor Gray
Write-Host "update_specific_slide_text 2 Chapter2 Machine_Learning Deep_learning_basics" -ForegroundColor Gray
Write-Host "save_presentation my_ai_presentation.pptx" -ForegroundColor Gray

Write-Host "`nSmart Text Update Features:" -ForegroundColor White
Write-Host "Finds Chapter text in template and replaces with chapter parameter" -ForegroundColor Gray
Write-Host "Finds Title text in template and replaces with title parameter" -ForegroundColor Gray
Write-Host "Finds Contents text in template and replaces with contents parameter" -ForegroundColor Gray
Write-Host "Auto font setting: Pretendard font with appropriate sizes" -ForegroundColor Gray

Write-Host "`nWorkflow:" -ForegroundColor White
Write-Host "1. Place Chapter Title Contents placeholder text in template" -ForegroundColor Gray
Write-Host "2. Use clone_template_and_update to replace all content at once" -ForegroundColor Gray
Write-Host "3. Save with save_presentation" -ForegroundColor Gray
Write-Host "4. Use update_specific_slide_text for individual slide modifications" -ForegroundColor Gray

Write-Host "`nSetup completed successfully!" -ForegroundColor Green
Write-Host "`nIMPORTANT: Copy the Python server code to $workDir\enhanced_ppt_server.py" -ForegroundColor Red
Write-Host "Then restart Claude Desktop and test with: scan_templates" -ForegroundColor Yellow