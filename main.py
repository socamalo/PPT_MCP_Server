from mcp.server.fastmcp import FastMCP
import win32com.client
import os
import uuid
from typing import Dict, List, Optional, Any

mcp = FastMCP("ppts")

USER_AGENT = "ppts-app/1.0"

class PPTAutomation:
    def __init__(self):
        self.ppt_app = None
        self.presentations = {}  # Store presentation IDs and their objects
        
    def initialize(self):
        try:
            # Try to connect to a running PowerPoint instance
            self.ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            return True
        except:
            try:
                # If no instance is running, create a new one
                self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                self.ppt_app.Visible = True
                return True
            except:
                return False
                
    def get_open_presentations(self):
        """Get all currently open presentations in PowerPoint"""
        result = []
        if not self.ppt_app:
            self.initialize()
            
        if self.ppt_app:
            for i in range(1, self.ppt_app.Presentations.Count + 1):
                pres = self.ppt_app.Presentations.Item(i)
                pres_id = str(uuid.uuid4())
                self.presentations[pres_id] = pres
                result.append({
                    "id": pres_id,
                    "name": os.path.basename(pres.FullName) if pres.FullName else "Untitled",
                    "path": pres.FullName,
                    "slide_count": pres.Slides.Count
                })
        return result

# Create a global instance of our automation class
ppt_automation = PPTAutomation()

@mcp.tool()
def initialize_powerpoint() -> bool:
    """Initialize connection to PowerPoint and make it visible if it wasn't already running."""
    return ppt_automation.initialize()

@mcp.tool()
def get_presentations() -> List[Dict[str, Any]]:
    """Get a list of all open PowerPoint presentations with their metadata."""
    return ppt_automation.get_open_presentations()

@mcp.tool()
def open_presentation(path: str) -> Dict[str, Any]:
    """
    Open a PowerPoint presentation from the specified path.
    
    Args:
        path: Full path to the PowerPoint file (.pptx, .ppt)
        
    Returns:
        Dictionary with presentation ID and metadata
    """
    if not ppt_automation.ppt_app:
        ppt_automation.initialize()
        
    if not os.path.exists(path):
        return {"error": f"File not found: {path}"}
    
    try:
        pres = ppt_automation.ppt_app.Presentations.Open(path)
        pres_id = str(uuid.uuid4())
        ppt_automation.presentations[pres_id] = pres
        
        return {
            "id": pres_id,
            "name": os.path.basename(path),
            "path": path,
            "slide_count": pres.Slides.Count
        }
    except Exception as e:
        return {"error": str(e)}

@mcp.tool()
def get_slides(presentation_id: str) -> List[Dict[str, Any]]:
    """
    Get a list of all slides in a presentation.
    
    Args:
        presentation_id: ID of the presentation
        
    Returns:
        List of slide metadata
    """
    if presentation_id not in ppt_automation.presentations:
        return {"error": "Presentation ID not found"}
    
    pres = ppt_automation.presentations[presentation_id]
    slides = []
    
    for i in range(1, pres.Slides.Count + 1):
        slide = pres.Slides.Item(i)
        slide_id = str(i)  # Using slide index as ID for simplicity
        
        slides.append({
            "id": slide_id,
            "index": i,
            "title": get_slide_title(slide),
            "shape_count": slide.Shapes.Count
        })
    
    return slides

def get_slide_title(slide):
    """Helper function to extract slide title if available"""
    try:
        # First check if there's a title placeholder
        for shape in slide.Shapes:
            if shape.Type == 14:  # msoPlaceholder
                if shape.PlaceholderFormat.Type == 1:  # ppPlaceholderTitle
                    if hasattr(shape, "TextFrame") and hasattr(shape.TextFrame, "TextRange"):
                        return shape.TextFrame.TextRange.Text
        
        # If no title placeholder found, check any shape with text
        for shape in slide.Shapes:
            if hasattr(shape, "TextFrame") and hasattr(shape.TextFrame, "TextRange"):
                try:
                    text = shape.TextFrame.TextRange.Text
                    if text and text.strip():
                        return text  # Return the first non-empty text as title
                except:
                    continue
    except:
        pass
    
    return "Untitled Slide"

@mcp.tool()
def get_slide_text(presentation_id: str, slide_id: int) -> Dict[str, Any]:
    """
    Get all text content in a slide.
    
    Args:
        presentation_id: ID of the presentation
        slide_id: ID of the slide (integer)
        
    Returns:
        Dictionary containing text content organized by shape
    """
    try:
        # Check if presentation exists
        if presentation_id not in ppt_automation.presentations:
            return {"error": f"Presentation ID not found: {presentation_id}"}
        
        pres = ppt_automation.presentations[presentation_id]
        
        # Get slide count
        try:
            slide_count = pres.Slides.Count
        except Exception as e:
            return {"error": f"Unable to get slide count: {str(e)}"}
            
        if slide_count == 0:
            return {"error": "Presentation has no slides"}
            
        # Check slide_id range
        if slide_id < 1 or slide_id > slide_count:
            return {"error": f"Invalid slide ID: {slide_id}. Valid range is 1-{slide_count}"}
        
        # Safely get the slide
        try:
            slide = pres.Slides.Item(int(slide_id))
        except Exception as e:
            return {"error": f"Error retrieving slide: {str(e)}"}
        
        text_content = {}
        
        # Process all shapes on the slide
        shape_count = 0
        try:
            shape_count = slide.Shapes.Count
        except Exception as e:
            return {"error": f"Unable to get shape count: {str(e)}"}
            
        for shape_idx in range(1, shape_count + 1):
            try:
                shape = slide.Shapes.Item(shape_idx)
                shape_id = str(shape_idx)
                
                # Check if the shape has a text frame
                has_text = False
                text = ""
                
                try:
                    # First try TextFrame2 (PowerPoint 2010 and higher)
                    if hasattr(shape, "TextFrame2") and shape.TextFrame2.HasText:
                        has_text = True
                        text = shape.TextFrame2.TextRange.Text
                    # Then try older TextFrame
                    elif hasattr(shape, "TextFrame") and hasattr(shape.TextFrame, "HasText") and shape.TextFrame.HasText:
                        has_text = True
                        text = shape.TextFrame.TextRange.Text
                    elif hasattr(shape, "TextFrame") and hasattr(shape.TextFrame, "TextRange"):
                        try:
                            text = shape.TextFrame.TextRange.Text
                            has_text = bool(text and text.strip())
                        except:
                            pass
                except Exception as shape_text_error:
                    continue  # Skip this shape if text cannot be retrieved
                
                if has_text or (text and text.strip()):
                    shape_name = "Unnamed Shape"
                    try:
                        shape_name = shape.Name
                    except:
                        pass
                        
                    text_content[shape_id] = {
                        "shape_name": shape_name,
                        "text": text
                    }
            except Exception as shape_error:
                continue  # Skip this shape if an error occurs without interrupting the process
        
        return {
            "slide_id": slide_id,
            "slide_index": slide_id,
            "slide_count": slide_count,
            "shape_count": shape_count,
            "content": text_content
        }
    except Exception as e:
        # Catch all other exceptions
        return {
            "error": f"An error occurred: {str(e)}",
            "presentation_id": presentation_id,
            "slide_id": slide_id
        }

@mcp.tool()
def update_text(presentation_id: str, slide_id: str, shape_id: str, text: str) -> Dict[str, Any]:
    """
    Update the text content of a shape.
    
    Args:
        presentation_id: ID of the presentation
        slide_id: ID of the slide (numeric string)
        shape_id: ID of the shape (numeric string)
        text: New text content
        
    Returns:
        Status of the operation
    """
    if presentation_id not in ppt_automation.presentations:
        return {"error": "Presentation ID not found"}
    
    pres = ppt_automation.presentations[presentation_id]
    
    # Properly handle string inputs that might contain quotes
    try:
        # Remove quotes if they exist in the strings
        clean_slide_id = slide_id.strip('"\'')
        clean_shape_id = shape_id.strip('"\'')
        
        slide_idx = int(clean_slide_id)
        shape_idx = int(clean_shape_id)
    except ValueError as e:
        return {"error": f"Invalid ID format: {str(e)}"}
    
    if slide_idx < 1 or slide_idx > pres.Slides.Count:
        return {"error": f"Invalid slide ID: {slide_id}"}
    
    try:
        slide = pres.Slides.Item(slide_idx)
    except Exception as e:
        return {"error": f"Error accessing slide: {str(e)}"}
    
    if shape_idx < 1 or shape_idx > slide.Shapes.Count:
        return {"error": f"Invalid shape ID: {shape_id}"}
    
    try:
        shape = slide.Shapes.Item(shape_idx)
        
        # First try TextFrame2 (newer PowerPoint versions)
        if hasattr(shape, "TextFrame2") and shape.TextFrame2.HasText:
            shape.TextFrame2.TextRange.Text = text
            return {"success": True, "message": "Text updated successfully using TextFrame2"}
        
        # Then try TextFrame
        elif hasattr(shape, "TextFrame") and hasattr(shape.TextFrame, "TextRange"):
            shape.TextFrame.TextRange.Text = text
            return {"success": True, "message": "Text updated successfully using TextFrame"}
            
        # Try finding text in grouped shapes
        elif shape.Type == 6:  # msoGroup (grouped shapes)
            updated = False
            for i in range(1, shape.GroupItems.Count + 1):
                subshape = shape.GroupItems.Item(i)
                if hasattr(subshape, "TextFrame") and hasattr(subshape.TextFrame, "TextRange"):
                    subshape.TextFrame.TextRange.Text = text
                    updated = True
                    break
                elif hasattr(subshape, "TextFrame2") and subshape.TextFrame2.HasText:
                    subshape.TextFrame2.TextRange.Text = text
                    updated = True
                    break
            
            if updated:
                return {"success": True, "message": "Text updated successfully in grouped shape"}
            else:
                return {"success": False, "message": "No text frame found in grouped shape"}
                
        else:
            return {"success": False, "message": "Shape does not contain editable text"}
    except Exception as e:
        return {"success": False, "error": f"Error updating text: {str(e)}"}

@mcp.tool()
def save_presentation(presentation_id: str, path: str = None) -> Dict[str, Any]:
    """
    Save a presentation to disk.
    
    Args:
        presentation_id: ID of the presentation
        path: Optional path to save the file (if None, save to current location)
        
    Returns:
        Status of the operation
    """
    if presentation_id not in ppt_automation.presentations:
        return {"error": "Presentation ID not found"}
    
    pres = ppt_automation.presentations[presentation_id]
    
    try:
        if path:
            pres.SaveAs(path)
        else:
            pres.Save()
        return {
            "success": True, 
            "path": path if path else pres.FullName
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

@mcp.tool()
def close_presentation(presentation_id: str, save: bool = True) -> Dict[str, Any]:
    """
    Close a presentation.
    
    Args:
        presentation_id: ID of the presentation
        save: Whether to save changes before closing
        
    Returns:
        Status of the operation
    """
    if presentation_id not in ppt_automation.presentations:
        return {"error": "Presentation ID not found"}
    
    pres = ppt_automation.presentations[presentation_id]
    
    try:
        pres.Close(save)
        del ppt_automation.presentations[presentation_id]
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}

@mcp.tool()
def create_presentation() -> Dict[str, Any]:
    """
    Create a new PowerPoint presentation.
    
    Returns:
        Dictionary containing new presentation ID and metadata
    """
    if not ppt_automation.ppt_app:
        ppt_automation.initialize()
        
    try:
        pres = ppt_automation.ppt_app.Presentations.Add()
        pres_id = str(uuid.uuid4())
        ppt_automation.presentations[pres_id] = pres
        
        return {
            "id": pres_id,
            "name": "New Presentation",
            "path": "",
            "slide_count": pres.Slides.Count
        }
    except Exception as e:
        return {"error": str(e)}

@mcp.tool()
def add_slide(presentation_id: str, layout_type: int = 1) -> Dict[str, Any]:
    """
    Add a new slide to the presentation.
    
    Args:
        presentation_id: ID of the presentation
        layout_type: Slide layout type (default is 1, title slide)
            1: ppLayoutTitle (title slide)
            2: ppLayoutText (slide with title and text)
            3: ppLayoutTwoColumns (two-column slide)
            7: ppLayoutBlank (blank slide)
            etc...
            
    Returns:
        Information about the new slide
    """
    if presentation_id not in ppt_automation.presentations:
        return {"error": "Presentation ID not found"}
    
    pres = ppt_automation.presentations[presentation_id]
    
    try:
        # Get current slide count
        slide_index = pres.Slides.Count + 1
        
        # Add new slide
        slide = pres.Slides.Add(slide_index, layout_type)
        
        return {
            "id": str(slide_index),
            "index": slide_index,
            "title": "New Slide",
            "shape_count": slide.Shapes.Count
        }
    except Exception as e:
        return {"error": f"Error adding slide: {str(e)}"}

@mcp.tool()
def add_text_box(presentation_id: str, slide_id: str, text: str, 
                 left: float = 100, top: float = 100, 
                 width: float = 400, height: float = 200) -> Dict[str, Any]:
    """
    Add a text box to a slide and set its text content.
    
    Args:
        presentation_id: ID of the presentation
        slide_id: ID of the slide (numeric string)
        text: Text content
        left: Left edge position of the text box (points)
        top: Top edge position of the text box (points)
        width: Width of the text box (points)
        height: Height of the text box (points)
        
    Returns:
        Operation status and ID of the new shape
    """
    if presentation_id not in ppt_automation.presentations:
        return {"error": "Presentation ID not found"}
    
    pres = ppt_automation.presentations[presentation_id]
    
    try:
        # Ensure slide_id is an integer
        slide_idx = int(slide_id.strip('"\''))
        
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            return {"error": f"Invalid slide ID: {slide_id}"}
        
        slide = pres.Slides.Item(slide_idx)
        
        # Add text box
        shape = slide.Shapes.AddTextbox(1, left, top, width, height)  # 1 = msoTextOrientationHorizontal
        
        # Set text content
        shape.TextFrame.TextRange.Text = text
        
        # Get the new shape's index
        shape_id = None
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes.Item(i) == shape:
                shape_id = str(i)
                break
        
        return {
            "success": True,
            "slide_id": slide_id,
            "shape_id": shape_id,
            "message": "Text box added successfully"
        }
    except Exception as e:
        return {"error": f"Error adding text box: {str(e)}"}

@mcp.tool()
def set_slide_title(presentation_id: str, slide_id: str, title: str) -> Dict[str, Any]:
    """
    Set the title text of a slide.
    
    Args:
        presentation_id: ID of the presentation
        slide_id: ID of the slide (numeric string)
        title: New title text
        
    Returns:
        Status of the operation
    """
    if presentation_id not in ppt_automation.presentations:
        return {"error": "Presentation ID not found"}
    
    pres = ppt_automation.presentations[presentation_id]
    
    try:
        # Ensure slide_id is an integer
        slide_idx = int(slide_id.strip('"\''))
        
        if slide_idx < 1 or slide_idx > pres.Slides.Count:
            return {"error": f"Invalid slide ID: {slide_id}"}
        
        slide = pres.Slides.Item(slide_idx)
        
        # Find title placeholder
        title_found = False
        for shape in slide.Shapes:
            if shape.Type == 14:  # msoPlaceholder
                if hasattr(shape, "PlaceholderFormat") and shape.PlaceholderFormat.Type == 1:  # ppPlaceholderTitle
                    if hasattr(shape, "TextFrame") and hasattr(shape.TextFrame, "TextRange"):
                        shape.TextFrame.TextRange.Text = title
                        title_found = True
                        break
        
        if not title_found:
            # If no title placeholder found, add a text box as title
            shape = slide.Shapes.AddTextbox(1, 50, 50, 600, 50)
            shape.TextFrame.TextRange.Text = title
            
            # Set text format as title style
            shape.TextFrame.TextRange.Font.Size = 44
            shape.TextFrame.TextRange.Font.Bold = True
        
        return {
            "success": True,
            "message": "Slide title has been set"
        }
    except Exception as e:
        return {"error": f"Error setting slide title: {str(e)}"}

if __name__ == "__main__":
    mcp.run(transport="stdio")
