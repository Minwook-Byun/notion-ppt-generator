#!/usr/bin/env python3
"""
Enhanced PowerPoint MCP Server with Template Clone & Update
Core functionality: Template cloning and smart content update
"""

import os
import json
import datetime
import shutil
from pathlib import Path
from typing import Optional, Dict, List
from mcp.server.fastmcp import FastMCP
from pptx import Presentation
from pptx.util import Pt
import tempfile

# Create FastMCP server
mcp = FastMCP("Enhanced PowerPoint MCP Server with Template Clone & Update")

# Global presentation object
current_presentation = None
current_filename = None

# ═══════════════════════════════════════════════════════════════════
# TEMPLATE AND DIRECTORY CONFIGURATION
# ═══════════════════════════════════════════════════════════════════

# Common template paths (accessible to all users)
TEMPLATE_PATHS = {
    'common_templates': Path("C:/Templates/PowerPoint"),
    'public_documents': Path.home() / "Documents" / "PowerPoint Templates",
    'desktop_templates': Path.home() / "Desktop" / "Templates",
    'shared_drive': Path("//shared/templates/powerpoint"),
}

# Setup save directories
PRESENTATIONS_DIR = Path.home() / "Desktop" / "MyPPT"
PRESENTATIONS_DIR.mkdir(exist_ok=True)

# Setup temp directory
TEMP_DIR = Path(tempfile.gettempdir()) / "mcp_powerpoint"
TEMP_DIR.mkdir(exist_ok=True)

# Template registry
template_registry = {}

def discover_templates():
    """Discover available PowerPoint templates from common locations"""
    global template_registry
    template_registry = {}
    
    for location_name, path in TEMPLATE_PATHS.items():
        if path.exists():
            templates = list(path.glob("*.pptx")) + list(path.glob("*.potx"))
            for template_path in templates:
                template_name = template_path.stem
                template_registry[template_name] = {
                    'path': str(template_path),
                    'location': location_name,
                    'name': template_name,
                    'extension': template_path.suffix
                }
    
    return template_registry

def load_template_presentation(template_name: str) -> Optional[Presentation]:
    """Load a PowerPoint template by name"""
    if template_name not in template_registry:
        return None
    
    template_info = template_registry[template_name]
    template_path = template_info['path']
    
    try:
        return Presentation(template_path)
    except Exception as e:
        print(f"Error loading template {template_name}: {e}")
        return None

def update_presentation_with_smart_text(presentation: Presentation, chapter_text: str = "", 
                                      title_text: str = "", contents_text: str = "") -> Dict:
    """
    Smart text update for presentations.
    Finds placeholder text in template and replaces with actual content
    """
    modified_count = 0
    
    try:
        for slide_idx, slide in enumerate(presentation.slides):
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue

                text = shape.text_frame.text.strip().lower()

                # Chapter text update
                if any(keyword in text for keyword in ['chapter', 'chap']):
                    if chapter_text:
                        shape.text_frame.text = chapter_text
                        # Font settings
                        for paragraph in shape.text_frame.paragraphs:
                            paragraph.font.name = "Pretendard"
                            paragraph.font.size = Pt(18)
                        modified_count += 1
                
                # Title text update
                elif any(keyword in text for keyword in ['title']):
                    if title_text:
                        shape.text_frame.text = title_text
                        # Font settings
                        for paragraph in shape.text_frame.paragraphs:
                            paragraph.font.name = "Pretendard"
                            paragraph.font.size = Pt(24)
                        modified_count += 1

                # Contents text update
                elif any(keyword in text for keyword in ['contents', 'content']):
                    if contents_text:
                        shape.text_frame.text = contents_text
                        # Font settings
                        for paragraph in shape.text_frame.paragraphs:
                            paragraph.font.name = "Pretendard"
                            paragraph.font.size = Pt(14)
                        modified_count += 1
        
        return {
            "status": "success",
            "modified_count": modified_count,
            "error_message": None
        }

    except Exception as e:
        return {
            "status": "failure",
            "modified_count": 0,
            "error_message": str(e)
        }

# ═══════════════════════════════════════════════════════════════════
# TEMPLATE DISCOVERY AND MANAGEMENT TOOLS
# ═══════════════════════════════════════════════════════════════════

@mcp.tool()
def scan_templates() -> str:
    """Scan and discover available PowerPoint templates"""
    try:
        templates = discover_templates()
        
        if not templates:
            return """No templates found.
            
Searched locations:
""" + "\n".join([f"   • {name}: {path}" for name, path in TEMPLATE_PATHS.items()])
        
        result = [f"Found {len(templates)} templates:\n"]
        
        # Group by location
        by_location = {}
        for template_name, info in templates.items():
            location = info['location']
            if location not in by_location:
                by_location[location] = []
            by_location[location].append(info)
        
        for location, template_list in by_location.items():
            result.append(f"{location.replace('_', ' ').title()}:")
            for template in template_list:
                result.append(f"   {template['name']}{template['extension']}")
            result.append("")
        
        return "\n".join(result)
        
    except Exception as e:
        return f"Error scanning templates: {str(e)}"

@mcp.tool()
def list_available_templates() -> str:
    """List all currently available templates"""
    if not template_registry:
        discover_templates()
    
    if not template_registry:
        return "No templates available. Run 'scan_templates' first."
    
    result = [f"Available templates ({len(template_registry)}):\n"]
    
    for template_name, info in sorted(template_registry.items()):
        result.append(f"{template_name}")
        result.append(f"   Location: {info['location']}")
        result.append(f"   Path: {info['path']}")
        result.append("")
    
    return "\n".join(result)

# ═══════════════════════════════════════════════════════════════════
# CORE FEATURE: TEMPLATE CLONE AND CONTENT UPDATE
# ═══════════════════════════════════════════════════════════════════

@mcp.tool()
def clone_template_and_update(template_name: str, chapter: str = "", title: str = "", 
                             contents: str = "", new_filename: str = "") -> str:
    """
    CORE FEATURE: Clone template and update chapter, title, contents
    """
    global current_presentation, current_filename
    
    if not template_registry:
        discover_templates()
    
    if template_name not in template_registry:
        available = ", ".join(template_registry.keys()) if template_registry else "None"
        return f"Template '{template_name}' not found.\nAvailable templates: {available}"
    
    try:
        # 1. Load template
        current_presentation = load_template_presentation(template_name)
        
        if current_presentation is None:
            return f"Failed to load template '{template_name}'"
        
        # 2. Smart text update
        update_result = update_presentation_with_smart_text(
            current_presentation, chapter, title, contents
        )
        
        if update_result["status"] != "success":
            return f"Failed to update content: {update_result['error_message']}"
        
        # 3. Set filename
        if new_filename:
            current_filename = new_filename if new_filename.endswith('.pptx') else f"{new_filename}.pptx"
        else:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            current_filename = f"{template_name}_updated_{timestamp}.pptx"
        
        template_info = template_registry[template_name]
        
        return f"""Template cloned and updated successfully!

Template: {template_name}
Source: {template_info['location']}
Modified elements: {update_result['modified_count']}
Total slides: {len(current_presentation.slides)}

Updated content:
   Chapter: {chapter if chapter else '(unchanged)'}
   Title: {title if title else '(unchanged)'}
   Contents: {contents[:50] + '...' if len(contents) > 50 else contents if contents else '(unchanged)'}

Ready to save as: {current_filename}
Use 'save_presentation()' to save the updated presentation!"""
        
    except Exception as e:
        return f"Error cloning and updating template: {str(e)}"

@mcp.tool()
def update_specific_slide_text(slide_number: int, chapter: str = "", title: str = "", 
                              contents: str = "") -> str:
    """
    Update text in a specific slide
    """
    global current_presentation
    
    if current_presentation is None:
        return "No presentation open. Please create or load a presentation first."
    
    if slide_number < 1 or slide_number > len(current_presentation.slides):
        return f"Invalid slide number. Presentation has {len(current_presentation.slides)} slides."
    
    try:
        slide = current_presentation.slides[slide_number - 1]
        modified_count = 0
        
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            text = shape.text_frame.text.strip().lower()

            # Chapter text update
            if any(keyword in text for keyword in ['chapter', 'chap']) and chapter:
                shape.text_frame.text = chapter
                for paragraph in shape.text_frame.paragraphs:
                    paragraph.font.name = "Pretendard"
                    paragraph.font.size = Pt(18)
                modified_count += 1
            
            # Title text update
            elif any(keyword in text for keyword in ['title']) and title:
                shape.text_frame.text = title
                for paragraph in shape.text_frame.paragraphs:
                    paragraph.font.name = "Pretendard"
                    paragraph.font.size = Pt(24)
                modified_count += 1

            # Contents text update
            elif any(keyword in text for keyword in ['contents', 'content']) and contents:
                shape.text_frame.text = contents
                for paragraph in shape.text_frame.paragraphs:
                    paragraph.font.name = "Pretendard"
                    paragraph.font.size = Pt(14)
                modified_count += 1
        
        return f"""Slide {slide_number} updated successfully!

Modified elements: {modified_count}
Updated content:
   Chapter: {chapter if chapter else '(unchanged)'}
   Title: {title if title else '(unchanged)'}
   Contents: {contents[:50] + '...' if len(contents) > 50 else contents if contents else '(unchanged)'}"""
        
    except Exception as e:
        return f"Error updating slide {slide_number}: {str(e)}"

@mcp.tool()
def create_presentation_from_template(template_name: str, presentation_title: str = "New Presentation") -> str:
    """Create a new presentation from a template"""
    global current_presentation, current_filename
    
    if not template_registry:
        discover_templates()
    
    if template_name not in template_registry:
        available = ", ".join(template_registry.keys()) if template_registry else "None"
        return f"Template '{template_name}' not found.\nAvailable templates: {available}"
    
    try:
        current_presentation = load_template_presentation(template_name)
        
        if current_presentation is None:
            return f"Failed to load template '{template_name}'"
        
        current_filename = None
        
        # Update title slide if it exists
        if len(current_presentation.slides) > 0:
            title_slide = current_presentation.slides[0]
            if title_slide.shapes.title:
                title_slide.shapes.title.text = presentation_title
        
        template_info = template_registry[template_name]
        
        return f"""Presentation created from template!
Template: {template_name}
Source: {template_info['location']}
Title: {presentation_title}
Initial slides: {len(current_presentation.slides)}"""
        
    except Exception as e:
        return f"Error creating presentation from template: {str(e)}"

# ═══════════════════════════════════════════════════════════════════
# FILE MANAGEMENT TOOLS
# ═══════════════════════════════════════════════════════════════════

@mcp.tool()
def save_presentation(filename: Optional[str] = None, auto_save: bool = True) -> str:
    """Save presentation to local storage"""
    global current_presentation, current_filename
    
    if current_presentation is None:
        return "No presentation to save."
    
    try:
        if filename is None:
            if current_filename:
                filename = current_filename
            else:
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"presentation_{timestamp}.pptx"
        
        if not filename.endswith('.pptx'):
            filename += '.pptx'
        
        save_path = PRESENTATIONS_DIR / filename
        
        if save_path.exists():
            backup_name = f"{save_path.stem}_backup_{datetime.datetime.now().strftime('%H%M%S')}.pptx"
            backup_path = PRESENTATIONS_DIR / backup_name
            shutil.copy2(save_path, backup_path)
        
        current_presentation.save(str(save_path))
        current_filename = filename
        
        save_info = {
            "filename": filename,
            "path": str(save_path),
            "size": save_path.stat().st_size,
            "saved_at": datetime.datetime.now().isoformat(),
            "slide_count": len(current_presentation.slides)
        }
        
        if auto_save:
            meta_path = PRESENTATIONS_DIR / f"{save_path.stem}_meta.json"
            with open(meta_path, 'w', encoding='utf-8') as f:
                json.dump(save_info, f, indent=2, ensure_ascii=False)
        
        return f"""Presentation saved successfully!
Filename: {filename}
Path: {save_path}
Size: {save_info['size']:,} bytes
Slides: {save_info['slide_count']}
Saved at: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"""
        
    except Exception as e:
        return f"Save error: {str(e)}"

@mcp.tool()
def get_presentation_info() -> str:
    """Get current presentation information"""
    global current_presentation, current_filename
    
    if current_presentation is None:
        return "No presentation open."
    
    try:
        slide_count = len(current_presentation.slides)
        filename = current_filename or "Not saved"
        
        slide_titles = []
        for i, slide in enumerate(current_presentation.slides, 1):
            title = "No title"
            if slide.shapes.title and slide.shapes.title.text:
                title = slide.shapes.title.text[:50]
            slide_titles.append(f"  {i}. {title}")
        
        return f"""Current presentation info:
Filename: {filename}
Slides: {slide_count}
Save directory: {PRESENTATIONS_DIR}
Status: Ready
Slide list:
{chr(10).join(slide_titles)}"""
        
    except Exception as e:
        return f"Info error: {str(e)}"

@mcp.tool()
def list_saved_presentations() -> str:
    """List all saved presentations"""
    try:
        pptx_files = list(PRESENTATIONS_DIR.glob("*.pptx"))
        
        if not pptx_files:
            return f"No saved presentations found.\nSave path: {PRESENTATIONS_DIR}"
        
        file_list = []
        for file_path in sorted(pptx_files, key=lambda x: x.stat().st_mtime, reverse=True):
            stat = file_path.stat()
            size_mb = stat.st_size / (1024 * 1024)
            modified = datetime.datetime.fromtimestamp(stat.st_mtime)
            
            file_list.append(f"{file_path.name}")
            file_list.append(f"   Size: {size_mb:.1f}MB")
            file_list.append(f"   Modified: {modified.strftime('%Y-%m-%d %H:%M:%S')}")
            file_list.append("")
        
        return f"""Saved presentations ({len(pptx_files)} files)
Save path: {PRESENTATIONS_DIR}
{chr(10).join(file_list)}"""
        
    except Exception as e:
        return f"Error listing files: {str(e)}"

@mcp.tool()
def create_presentation(title: str = "New Presentation") -> str:
    """Create a new PowerPoint presentation"""
    global current_presentation, current_filename
    
    try:
        current_presentation = Presentation()
        current_filename = None
        
        slide_layout = current_presentation.slide_layouts[0]
        slide = current_presentation.slides.add_slide(slide_layout)
        
        if slide.shapes.title:
            slide.shapes.title.text = title
            
        return f"New presentation created: '{title}'"
    except Exception as e:
        return f"Error creating presentation: {str(e)}"

@mcp.tool()
def add_slide(title: str = "", content: str = "", layout_index: int = 1) -> str:
    """Add new slide to presentation"""
    global current_presentation
    
    if current_presentation is None:
        return "No presentation open. Please create a presentation first."
    
    try:
        slide_layout = current_presentation.slide_layouts[layout_index]
        slide = current_presentation.slides.add_slide(slide_layout)
        
        if title and slide.shapes.title:
            slide.shapes.title.text = title
            
        if content and len(slide.placeholders) > 1:
            slide.placeholders[1].text = content
            
        slide_number = len(current_presentation.slides)
        return f"Added slide {slide_number}: '{title}'"
    except Exception as e:
        return f"Error adding slide: {str(e)}"

# Initialize template discovery on startup
discover_templates()

if __name__ == "__main__":
    print(f"Enhanced PowerPoint MCP Server with Template Clone & Update starting")
    print(f"Save directory: {PRESENTATIONS_DIR}")
    print(f"Temp directory: {TEMP_DIR}")
    print(f"Template registry: {len(template_registry)} templates found")
    mcp.run()