import logging
import sys
import os
from typing import Any,List,Dict,Optional

from mcp.server.fastmcp import FastMCP

from .exceptions import(
    PresentationError,
    SlideError,
    ShapeError,
    ChartError,
    SmartArtError,
    TableError,
    ConversionError
)

from .presentation import get_or_create_presentation
from .slide import append_slide as create_slide_impl
from .shape import add_shape as add_shape_impl
from .chart import add_chart as add_chart_impl
from .smartart import create_smartart as create_smartart_impl
from .table import create_table as create_table_impl
from .conversion import convert_presentation as convert_presentation_impl

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("spire-ppt-mcp.log")
    ],
    force=True
)

logger = logging.getLogger("spire-ppt-mcp")

# Get Ppt files path from environment or use default
PPT_FILES_PATH = os.environ.get("PPT_FILES_PATH", "./ppt_files")

# # Create the directory if it doesn't exist
# os.makedirs(PPT_FILES_PATH, exist_ok=True)

# Initialize FastMCP server
mcp = FastMCP(
    "spire-ppt-mcp",
    version="0.1.1",
    description="Spire.Ppt MCP Server for manipulating Ppt files",
    dependencies=["spire_presentation_free>=9.12.0"],
    env_vars={
        "PPT_FILES_PATH": {
            "description": "Path to Ppt files directory",
            "required": False,
            "default": PPT_FILES_PATH
        }
    }
)

def get_ppt_path(filename: str) -> str:
    """Get full path to Ppt file.
    
    Args:
        filename: Name of Ppt file
        
    Returns:
        Full path to Ppt file
    """
    # If filename is already an absolute path, return it
    if os.path.isabs(filename):
        return filename

    # Use the configured Ppt files path
    return os.path.join(PPT_FILES_PATH, filename)

@mcp.tool()
def create_presentation(filepath:str) -> str:
    """
    Creates a new Ppt presentation.
    
    Parameters:
    filepath (str): Path where the new presentation will be saved
    
    Returns:
    str: Success message with the created presentation path
    """
    try:
        full_path = get_ppt_path(filepath)
        result = get_or_create_presentation(full_path)
        return f"Created presentation at {full_path}"
    except PresentationError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error creating presentation:{e}")
        raise

@mcp.tool()
def create_slide(filepath:str) -> str:
    """
    Creates a new slide in an existing presentaion.
    
    Parameters:
    filepath (str): Path to the Ppt file
    
    Returns:
    str: Success message confirming slide creation
    """
    try:
        full_path = get_ppt_path(filepath)
        result = create_slide_impl(full_path)
        return str(result)
    except SlideError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error creating slide:{e}")
        raise

@mcp.tool()
def delete_slide(filepath:str,slide_num:int) -> dict[str,Any]:
    """
    Deletes a slide from an existing presentation.

    Parameters:
    filepath (str): Path to the PPT file
    slide_num (int): Index of the slide to delete (starting from 0)

    Returns:
    str: Success message confirming slide deletion
    """
    try:
        full_path = get_ppt_path(filepath)
        from .slide import delete_slide as delete_slide_impl
        result = delete_slide_impl(full_path,slide_num)
        return result
    except SlideError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error delete slide:{e}")
        raise

@mcp.tool()
def add_shape(
    filepath:str,
    slide_num:int = 0,
    x:float = 0,
    y:float = 0,
    width:float = 200,
    height:float = 200,
    shape_type:str = "Rectangle",
    line_color:str = None,
    fill_color:str = None
) -> dict[str,Any]:
    """
    Adds a geometric shape to a specified slide in a PowerPoint presentation.

    Parameters:
        filepath (str): Path to the PowerPoint file (.pptx).
        slide_num (int, optional): Slide number where the shape will be added.
                                   Defaults to 0 (typically the first slide).
        x (float, optional): X-coordinate of the shape's top-left corner in pixels. Defaults to 0.
        y (float, optional): Y-coordinate of the shape's top-left corner in pixels. Defaults to 0.
        width (float, optional): Width of the shape in pixels. Defaults to 200.
        height (float, optional): Height of the shape in pixels. Defaults to 200.
        shape_type (str, optional): Type of shape to add. Supported types depend on the underlying library.
                                    Common options include: 'Rectangle', 'Oval', 'Arrow', 'TextBox', etc.
                                    Defaults to 'Rectangle'.
        line_color (str, optional): Outline (stroke) color of the shape in hex format (e.g., '#FF5733').
                                    If not provided, the default outline style is used or no outline may be applied.
        fill_color (str, optional): Fill color of the shape in hex format (e.g., '#C0C0C0').
                                    If not provided, the shape will have no fill.

    Returns:
        Dict[str, Any]: A dictionary containing operation result with the following keys:
            - success (bool): Whether the operation was successful.
            - message (str): Detailed message or error description.
            - shape_info (dict, optional): Information about the added shape, such as:
                                           - type (str)
                                           - position (x, y)
                                           - size (width, height)
                                           - colors (line and fill if applicable)

    Raises:
        ShapeError: If adding the shape fails due to:
                    - Invalid parameters (e.g., invalid slide number or unsupported shape type)
                    - File access issues
                    - Unsupported operations by the underlying library
    """
    try:
        full_path = get_ppt_path(filepath)
        result = add_shape_impl(
            filepath=full_path,
            slide_num=slide_num,
            x = x,
            y = y,
            width=width,
            height=height,
            shape_type=shape_type,
            line_color=line_color,
            fill_color=fill_color
        )
        return result
    except ShapeError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error add shape:{e}")
        raise

@mcp.tool()
def delete_shape(filepath:str,slide_num:int,shape_num:int) -> dict[str,Any]:
    """
    Deletes a shape from a specified slide in a PowerPoint presentation.

    Parameters:
        filepath (str): Path to the PowerPoint file.
        slide_num (int): Slide number from which the shape will be deleted.
        shape_num (int): The index or identifier of the shape to delete. 
                         This depends on how shapes are indexed in the target library.

    Returns:
        Dict[str, Any]: A dictionary containing operation result with the following keys:
            - success (bool): Whether the operation was successful.
            - message (str): Detailed message or error description.
            - shape_info (dict, optional): Information about the deleted shape (if available and applicable).

    Raises:
        ShapeError: If deleting the shape fails due to invalid parameters, such as incorrect slide number or shape identifier,
                    or other issues like file access permissions.
    """
    try:
        full_path = get_ppt_path(filepath)
        from .shape import delete_shape as delete_shape_impl
        result = delete_shape_impl(full_path,slide_num,shape_num)
        return result
    except ShapeError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error delete shape:{e}")
        raise

@mcp.tool()
def add_text_shape(filepath:str,slide_num:int,shape_num:int = None,text:str = "") -> dict[str,Any]:
    """
    Adds a new text shape or updates an existing one on a specified slide in a PowerPoint presentation.

    Parameters:
        filepath (str): Path to the PowerPoint file (.pptx).
        slide_num (int): Slide number where the text shape will be added or updated.
                         The index typically starts at 0 or 1 depending on the library used.
        shape_num (int, optional): Index or identifier of an existing shape to update.
                                   If None, a new text shape will be created. Defaults to None.
        text (str, optional): Text content to insert into the shape. Defaults to an empty string.

    Returns:
        Dict[str, Any]: A dictionary containing operation result with the following keys:
            - success (bool): Whether the operation was successful.
            - message (str): Detailed message or error description.
            - shape_info (dict, optional): Information about the added or updated text shape,
                                           such as its index, position, size, and text content.

    Raises:
        ShapeError: If adding or updating the text shape fails due to:
                    - Invalid parameters (e.g., invalid slide number)
                    - File access issues
                    - Shape not found when trying to update
                    - Unsupported operations by the underlying library
    """
    try:
        full_path = get_ppt_path(filepath)
        from .shape import add_text_shape as add_text_shape_impl
        result = add_text_shape_impl(full_path,slide_num,shape_num,text)
        return result
    except ShapeError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error add text shape:{e}")
        raise

@mcp.tool()
def add_chart(
    filepath:str,
    slide_num:int,
    x:float = 0,
    y:float = 0,
    width:float = 200,
    height:float = 200,
    chart_type:str = "Pie"
) -> dict[str,Any]:
    """
    Adds a new chart to a specified location on a slide in a PowerPoint presentation.

    Parameters:
        filepath (str): Path to the PowerPoint file (.pptx).
        slide_num (int): Slide number where the chart will be added.
                         The index typically starts at 0 or 1 depending on the library used.
        x (float, optional): X-coordinate of the chart's top-left corner in pixels. Defaults to 0.
        y (float, optional): Y-coordinate of the chart's top-left corner in pixels. Defaults to 0.
        width (float, optional): Width of the chart in pixels. Defaults to 200.
        height (float, optional): Height of the chart in pixels. Defaults to 200.
        chart_type (str, optional): Type of chart to add. Supported types depend on the underlying library.
                                    Common options include: 'Bar', 'Column', 'Pie', 'Line', 'Scatter', etc.
                                    Defaults to 'Pie'.

    Returns:
        Dict[str, Any]: A dictionary containing operation result with the following keys:
            - success (bool): Whether the operation was successful.
            - message (str): Detailed message or error description.
            - chart_info (dict, optional): Information about the added chart, such as:
                                           - chart_type (str)
                                           - position (x, y)
                                           - size (width, height)
                                           - reference ID or other metadata (if applicable)

    Raises:
        ChartError: If adding the chart fails due to:
                    - Invalid parameters (e.g., invalid slide number or unsupported chart type)
                    - File access issues
                    - Unsupported operations by the underlying library
    """
    try:
        full_path = get_ppt_path(filepath)
        result = add_chart_impl(
            filepath=full_path,
            slide_num=slide_num,
            x = x,
            y = y,
            width=width,
            height=height,
            chart_type=chart_type
        )
        return result
    except  ChartError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error add chart:{e}")
        raise

@mcp.tool()
def create_smartart(
    filepath:str,
    slide_num:int,
    x:float = 0,
    y:float = 0,
    width:float = 200,
    height:float = 200,
    layout_type:str = "Gear"
) -> dict[str,Any]:
    """
    Adds a SmartArt graphic to a specified location on a slide in a PowerPoint presentation.

    Parameters:
        filepath (str): Path to the PowerPoint file (.pptx).
        slide_num (int): Slide number where the SmartArt will be added.
                         The index typically starts at 0 or 1 depending on the library used.
        x (float, optional): X-coordinate of the SmartArt's top-left corner in pixels. Defaults to 0.
        y (float, optional): Y-coordinate of the SmartArt's top-left corner in pixels. Defaults to 0.
        width (float, optional): Width of the SmartArt graphic in pixels. Defaults to 200.
        height (float, optional): Height of the SmartArt graphic in pixels. Defaults to 200.
        layout_type (str, optional): Type of SmartArt layout to add. Supported types depend on the underlying library.
                                     Common examples include: 'Gear', 'Bubbles', 'List', 'Process', 'Hierarchy', etc.
                                     Defaults to 'Gear'.

    Returns:
        Dict[str, Any]: A dictionary containing operation result with the following keys:
            - success (bool): Whether the operation was successful.
            - message (str): Detailed message or error description.
            - smartart_info (dict, optional): Information about the added SmartArt graphic, such as:
                                              - layout_type (str)
                                              - position (x, y)
                                              - size (width, height)
                                              - reference ID or other metadata (if applicable)

    Raises:
        SmartArtError: If creating or inserting the SmartArt fails due to:
                       - Invalid parameters (e.g., invalid slide number or unsupported layout type)
                       - File access issues
                       - Unsupported operations by the underlying library
    """
    try:
        full_path = get_ppt_path(filepath)
        result = create_smartart_impl(
            filepath=full_path,
            slide_num=slide_num,
            x = x,
            y = y,
            width=width,
            height=height,
            layout_type=layout_type
        )
        return result
    except  SmartArtError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error add chart:{e}")
        raise

@mcp.tool()
def shape_to_image(
    filepath:str,
    slide_num:int,
    output_filepath:str
) -> dict[str,Any]:
    """
    Converts shapes on a specified slide of a PowerPoint presentation into an image file.

    This function renders the slide content (including all shapes) into a static image and saves it to disk.

    Parameters:
        filepath (str): Path to the input PowerPoint file (.pptx).
        slide_num (int): Slide number to convert into an image.
                         The index typically starts at 0 or 1 depending on the library or backend used.
        output_filepath (str): Full path where the output image will be saved (e.g., 'output/slide.png').
                               The file extension determines the image format (e.g., .png, .jpg).

    Returns:
        Dict[str, Any]: A dictionary containing operation result with the following keys:
            - success (bool): Whether the operation was successful.
            - message (str): Detailed message or error description.
            - image_info (dict, optional): Additional information about the generated image, such as:
                                           - file_path (str): Path to the saved image.
                                           - dimensions (tuple): Width and height of the image.
                                           - slide_number (int): Index of the exported slide.

    Raises:
        ShapeError: If exporting the slide as an image fails due to:
                          - Invalid parameters (e.g., non-existent slide number)
                          - File access issues (e.g., permission denied)
                          - Backend rendering errors
    """
    try:
        full_path = get_ppt_path(filepath)
        from .shape import shape_to_image as shape_to_image_impl
        result = shape_to_image_impl(
            filepath=full_path,
            slide_num=slide_num,
            output_filepath = output_filepath
        )
        return result
    except ShapeError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error:{e}")
        raise
    
@mcp.tool()
def create_table(
    filepath:str,
    slide_num:int,
    x:float = 0,
    y:float = 0,
    widths:List[float] = [50,50],
    heights:List[float] = [20,20]
) -> dict[str,Any]:
    """
    Creates a table on a specified slide in a PowerPoint presentation.

    Parameters:
        filepath (str): Path to the PowerPoint file (.pptx).
        slide_num (int): Index of the slide where the table will be added.
                         The index typically starts at 0 or 1 depending on the library used.
        x (float, optional): X-coordinate of the table's top-left corner in pixels. Defaults to 0.
        y (float, optional): Y-coordinate of the table's top-left corner in pixels. Defaults to 0.
        widths (List[float], optional): A list specifying the width of each column in pixels. 
                                        Defaults to [50, 50] (two columns, 50 pixels wide each).
        heights (List[float], optional): A list specifying the height of each row in pixels.
                                         Defaults to [20, 20] (two rows, 20 pixels high each).

    Returns:
        Dict[str, Any]: A dictionary containing the result of the operation with the following keys:
            - success (bool): Whether the table was successfully created.
            - message (str): A detailed message or error description.
            - table_info (dict, optional): Information about the created table, such as:
                                           - position (x, y)
                                           - dimensions (number of rows and columns)
                                           - total_width (float): Total width of the table.
                                           - total_height (float): Total height of the table.
                                           - column_widths (List[float])
                                           - row_heights (List[float])

    Raises:
        TableError: If creating the table fails due to:
                    - Invalid file path
                    - Invalid slide number
                    - Empty or invalid widths/heights lists
                    - Unsupported operations by the underlying library
    """
    try:
        full_path = get_ppt_path(filepath)
        result = create_table_impl(
            filepath=full_path,
            slide_num=slide_num,
            x = x,
            y = y,
            widths = widths,
            heights = heights
        )
        return result
    except TableError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error:{e}")
        raise
    
@mcp.tool()
def add_text_table(
        filepath:str,
        slide_num:int,
        shape_num:int,
        data_str:List[str] = ["", "", "", ""]
) -> dict[str,Any]:
    """
    Adds or updates text content in a table shape on a specified slide of a PowerPoint presentation.

    The provided `data_str` list is used to populate the table cells row by row. 
    If the number of elements in `data_str` exceeds the number of table cells, extra elements are ignored.
    If there are fewer elements than cells, remaining cells are filled with empty strings.

    Parameters:
        filepath (str): Path to the PowerPoint file (.pptx).
        slide_num (int): Index of the slide containing the table (0-based index unless otherwise defined by the library).
        shape_num (int): Index of the shape on the slide; expected to be a table shape.
        data_str (List[str], optional): A list of strings used to fill the table cells row by row.
                                          Defaults to ["", "", "", ""].

    Returns:
        Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the operation was successful, False otherwise.
            - message (str): Description of the result or error.
            - table_info (dict, optional): Additional information about the updated table, such as:
                                           - rows (int): Number of rows in the table.
                                           - columns (int): Number of columns in the table.
                                           - filled_data (List[List[str]]): The 2D list used to fill the table.

    Raises:
        TableError: If the operation fails due to:
                    - Invalid file path or corrupted file
                    - Slide or shape index out of range
                    - Shape is not a table
                    - Library-specific errors when accessing or modifying the table
    """
    try:
        full_path = get_ppt_path(filepath)
        from .table import add_text_table as add_text_table_impl
        result = add_text_table_impl(
            filepath=full_path,
            slide_num=slide_num,
            shape_num = shape_num,
            data_str = data_str
        )
        return result
    except TableError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error:{e}")
        raise

@mcp.tool()
def set_shape_fill_picture(
    filepath:str,
    slide_num:int,
    shape_num:int,
    picture_url:str
) -> dict[str,Any]:
    """
    Sets the fill of a specified shape on a slide with an image from the given URL or local path.

    This function loads a PowerPoint presentation, accesses the specified slide and shape,
    and fills the shape with an image. The image can be provided via a local file path or URL.

    Parameters:
        filepath (str): Path to the input PowerPoint file (.pptx).
        slide_num (int): Index of the slide containing the target shape.
                         (Index typically starts at 0 depending on the library used)
        shape_num (int): Index of the shape within the slide to apply the image fill.
        picture_url (str): Path or URL to the image file. Supported formats typically include:
                           .png, .jpg, .gif, .bmp, etc., depending on backend support.

    Returns:
        Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the image was successfully applied, False otherwise.
            - message (str): Description of the result or error.
            - shape_info (dict, optional): Additional information about the updated shape, such as:
                                           - slide_number (int)
                                           - shape_index (int)
                                           - image_source (str): The original image path/URL used

    Raises:
        ShapeError: If applying the image fill fails due to:
                         - Invalid file path or corrupted PowerPoint file
                         - Slide or shape index out of range
                         - Shape does not support image fill
                         - Image file not found or unsupported format
                         - Backend-specific errors when modifying the shape
    """
    try:
        full_path = get_ppt_path(filepath)
        from .shape import fill_shape_with_picture as fill_shape_with_picture_impl
        result = fill_shape_with_picture_impl(
            filepath=full_path,
            slide_num=slide_num,
            shape_num = shape_num,
            picture_url = picture_url
        )
        return result
    except ShapeError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error:{e}")
        raise
    
@mcp.tool()
def get_shape_titles(
        filepath:str,
        output_filepath:str
) -> dict[str,Any]:
    """
    Extracts titles or text content from shapes in all slides of a PowerPoint presentation,
    and saves the extracted text into a plain text (.txt) file.

    This function loads the input presentation, iterates through all slides and shapes,
    extracts text (e.g., titles or labels) from applicable shapes, and writes that text
    into a new plain text file at the specified output path.

    Parameters:
        filepath (str): Path to the source PowerPoint file (.pptx).
        output_filepath (str): Path where the extracted shape texts will be saved as a .txt file.
                               The file extension should typically be .txt.

    Returns:
        Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the operation was successful, False otherwise.
            - message (str): Description of the result or error.
            - summary (dict, optional): Additional information such as:
                                        - total_slides (int): Number of slides processed.
                                        - total_shapes (int): Total number of shapes analyzed.
                                        - extracted_texts (List[str]): List of extracted texts from shapes.

    Raises:
        ShapeError: If the operation fails due to:
                              - Invalid or corrupted input file
                              - Permission issues writing to output file
                              - Unsupported shape types
                              - Library-specific errors during text extraction
    """
    try:
        full_path = get_ppt_path(filepath)
        from .shape import get_shape_titles as get_shape_titles_impl
        result = get_shape_titles_impl(
            filepath=full_path,
            output_filepath=output_filepath
        )
        return result
    except ShapeError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error:{e}")
        raise
            
@mcp.tool()
def group_shapes(
        filepath:str,
        slide_num:int,
        shape_num_list:List[int] = []
) -> dict[str,Any]:
    """
    Groups a list of shape objects on a specified slide into a single shape group within the PowerPoint presentation.

    This function loads the input presentation, accesses the specified slide,
    and groups the provided shape objects into a single group.
    
    Parameters:
        filepath (str): Path to the input PowerPoint file (.pptx).
        slide_num (int): Index of the slide containing the shapes to be grouped.
                         (Index typically starts at 0 depending on the library used)
        shape_num_list (List[int], optional): A list of indices of shapes on the slide to be grouped. Defaults to an empty list.

    Returns:
        Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the shapes were successfully grouped, False otherwise.
            - message (str): Description of the result or error.
            - group_info (dict, optional): Additional information about the grouping, such as:
                                           - slide_number (int)
                                           - grouped_shape_count (int): Number of shapes grouped

    Raises:
        ShapeError: If the operation fails due to:
                       - Invalid file path or corrupted PowerPoint file
                       - Slide index out of range
                       - Shape objects in `shape_list` are invalid or not on the same slide
                       - Not enough shapes selected for grouping (at least two required)
                       - Library-specific errors during grouping
    """
    try:
        full_path = get_ppt_path(filepath)
        from .shape import group_shapes as group_shapes_impl
        result = group_shapes_impl(
            filepath=full_path,
            slide_num=slide_num,
            shape_num_list = shape_num_list
        )
        return result
    except ShapeError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error:{e}")
        raise

@mcp.tool()
def ungroup_shapes(
        filepath:str,
        slide_num:int,
        shape_num:int
) -> dict[str,Any]:
    """
    Ungroups a grouped shape on a specified slide in a PowerPoint presentation.

    This function loads the input presentation, accesses the specified slide,
    and performs an ungroup operation on the shape at the given index, provided
    that the shape is a group shape.

    Parameters:
        filepath (str): Path to the input PowerPoint file (.pptx).
        slide_num (int): Index of the slide containing the group shape.
                         (Index typically starts at 0 depending on the library used)
        shape_num (int): Index of the shape on the slide that should be ungrouped.
    
    Returns:
        Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the shape was successfully ungrouped, False otherwise.
            - message (str): Description of the result or error.
            - ungroup_info (dict, optional): Additional information about the operation, such as:
                                             - slide_number (int)
                                             - ungrouped_shape_index (int)
                                             - number_of_ungrouped_shapes (int): Number of individual shapes released after ungrouping

    Raises:
        ShapeError: If the operation fails due to:
                         - Invalid file path or corrupted PowerPoint file
                         - Slide or shape index out of range
                         - The specified shape is not a group shape
                         - Library-specific errors during ungrouping
    """
    try:
        full_path = get_ppt_path(filepath)
        from .shape import ungroup_shapes as ungroup_shapes_impl
        result = ungroup_shapes_impl(
            filepath=full_path,
            slide_num=slide_num,
            shape_num = shape_num
        )
        return result
    except ShapeError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error:{e}")
        raise
    
@mcp.tool()
def change_slide_position(
    filepath: str,
    slide_num:int,
    slide_number:int
) -> dict[str,Any]:
    """
    Changes the position of a specified slide in the PowerPoint presentation to a new index.

    This function loads the input presentation and moves the slide at index `slide_num`
    to the new position specified by `slide_number`. All other slides are rearranged accordingly.

    Parameters:
        filepath (str): Path to the input PowerPoint file (.pptx).
        slide_num (int): Index of the slide to be moved (current position, 0-based or library-dependent).
        slide_number (int): Target index where the slide should be placed (0-based or library-dependent).

    Returns:
        Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the slide was successfully moved, False otherwise.
            - message (str): Description of the result or error.
            - move_info (dict, optional): Additional information about the slide movement, such as:
                                          - original_position (int)
                                          - new_position (int)

    Raises:
        SlideError: If the operation fails due to:
                        - Invalid file path or corrupted PowerPoint file
                        - Either `slide_num` or `slide_number` is out of range
                        - Attempting to move a slide to the same position
                        - Library-specific errors when reordering slides
    """
    try:
        full_path = get_ppt_path(filepath)
        from .slide import change_slide_position as change_slide_position_impl
        result = change_slide_position_impl(
            filepath=full_path,
            slide_num=slide_num,
            slide_number = slide_number
        )
        return result
    except SlideError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error:{e}")
        raise
    
@mcp.tool()
def add_image_in_master(
        filepath: str, 
        image_filepath: str,
        master_num:int = 0,
        x:float = 0,
        y:float = 0,
        width:float = 200,
        height:float = 200
) -> dict[str,Any]:
    """
    Adds an image to a specified slide master in a PowerPoint presentation.

    This function loads the input presentation, accesses the specified slide master,
    and inserts an image at the given position with the specified size.

    Parameters:
        filepath (str): Path to the input PowerPoint file (.pptx).
        image_filepath (str): Path to the image file to be inserted into the slide master.
                              Supported formats typically include: .png, .jpg, .gif, etc.
        master_num (int, optional): Index of the slide master to which the image will be added.
                                    Defaults to 0 (the first master).
        x (float, optional): X-coordinate (in points) for the top-left corner of the image.
                             Defaults to 0.
        y (float, optional): Y-coordinate (in points) for the top-left corner of the image.
                             Defaults to 0.
        width (float, optional): Width of the inserted image (in points). Defaults to 200.
        height (float, optional): Height of the inserted image (in points). Defaults to 200.

    Returns:
        Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the image was successfully added to the master, False otherwise.
            - message (str): Description of the result or error.
            - master_info (dict, optional): Additional information about the operation, such as:
                                            - master_index (int)
                                            - image_position (tuple[float, float]): (x, y)
                                            - image_size (tuple[float, float]): (width, height)

    Raises:
        SlideError: If the operation fails due to:
                          - Invalid file path or corrupted PowerPoint file
                          - Slide master index out of range
                          - Image file not found or unsupported format
                          - Library-specific errors when modifying the slide master
    """
    try:
        full_path = get_ppt_path(filepath)
        from .slide import add_image_in_master as add_image_in_master_impl
        result = add_image_in_master_impl(
            filepath=full_path,
            image_filepath=image_filepath,
            master_num = master_num,
            x = x,
            y = y,
            width = width,
            height = height
        )
        return result
    except SlideError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error:{e}")
        raise

@mcp.tool()
def set_alignment(
        filepath:str,
        slide_num:int = 0,
        shape_num:int = 0,
        paragraph_num:int = 0,
        text_alignment_type:str = "Left"
) -> dict[str,Any]:
    """
    Sets the text alignment for a specific paragraph within a shape on a given slide in a PowerPoint presentation.

    This function loads the input presentation, accesses the specified slide and shape,
    and applies the desired text alignment to the specified paragraph.

    Parameters:
        filepath (str): Path to the input PowerPoint file (.pptx).
        slide_num (int, optional): Index of the slide containing the target shape. Defaults to 0.
        shape_num (int, optional): Index of the shape within the slide that contains the target paragraph. Defaults to 0.
        paragraph_num (int, optional): Index of the paragraph within the shape whose alignment needs to be changed. Defaults to 0.
        text_alignment_type (str, optional): Desired text alignment type. Possible values are 'Left', 'Right', 'Center', 'Justify'.
                                             Defaults to 'Left'.

    Returns:
        Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the text alignment was successfully applied, False otherwise.
            - message (str): Description of the result or error.
            - alignment_info (dict, optional): Additional information about the operation, such as:
                                               - slide_number (int)
                                               - shape_index (int)
                                               - paragraph_index (int)
                                               - applied_alignment (str)

    Raises:
        ShapeError: If the operation fails due to:
                        - Invalid file path or corrupted PowerPoint file
                        - Slide, shape, or paragraph index out of range
                        - Unsupported text alignment type provided
                        - Library-specific errors during alignment setting
    """
    try:
        full_path = get_ppt_path(filepath)
        from .shape import set_alignment as set_alignment_impl
        result = set_alignment_impl(
            filepath=full_path,
            slide_num=slide_num,
            shape_num = shape_num,
            paragraph_num = paragraph_num,
            text_alignment_type = text_alignment_type
        )
        return result
    except ShapeError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error:{e}")
        raise

@mcp.tool()
def append_html(
        filepath:str,
        slide_num:int = 0,
        shape_num:int = 0,
        code_html:str = " "
) -> dict[str,Any]:
    """
    Appends HTML-formatted text content into a specified text frame within a shape on a given slide.

    This function loads the input PowerPoint presentation, accesses the specified slide and shape,
    and appends rich-text content defined by an HTML string into the shape's text frame.
    
    Note: The level of HTML formatting supported may vary depending on the underlying library used.
          Only basic tags (e.g., <b>, <i>, <u>, <font>, <p>) are typically supported.

    Parameters:
        filepath (str): Path to the input PowerPoint file (.pptx).
        slide_num (int, optional): Index of the slide containing the target shape. Defaults to 0.
        shape_num (int, optional): Index of the shape on the slide that contains the text frame. Defaults to 0.
        code_html (str, optional): HTML-formatted string to be appended into the text frame. 
                                   Defaults to a single space (" ").

    Returns:
        Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the HTML content was successfully appended, False otherwise.
            - message (str): Description of the result or error.
            - html_info (dict, optional): Additional information about the operation, such as:
                                           - slide_number (int)
                                           - shape_index (int)
                                           - applied_html (str)

    Raises:
        ShapeError: If the operation fails due to:
                         - Invalid file path or corrupted PowerPoint file
                         - Slide or shape index out of range
                         - Shape does not contain a valid text frame
                         - Unsupported or malformed HTML content
                         - Library-specific errors during HTML appending
    """
    try:
        full_path = get_ppt_path(filepath)
        from .shape import append_html as append_html_impl
        result = append_html_impl(
            filepath=full_path,
            slide_num=slide_num,
            shape_num = shape_num,
            code_html = code_html
        )
        return result
    except ShapeError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error:{e}")
        raise

@mcp.tool()
def set_autofittext(
        filepath:str,
        slide_num:int = 0,
        shape_num:int = 0,
        autofit_type:str = "Shape"
) -> dict[str,Any]:
    """
    Sets the auto-fit behavior for the text within a specified shape on a given slide in a PowerPoint presentation.

    This function loads the input presentation and configures how the text inside the specified shape adjusts 
    to fit within its bounding box. The auto-fit behavior can be applied to resize the text, shape, or none.

    Parameters:
        filepath (str): Path to the input PowerPoint file (.pptx).
        slide_num (int, optional): Index of the slide containing the target shape. Defaults to 0.
        shape_num (int, optional): Index of the shape on the slide whose text auto-fit behavior will be set. Defaults to 0.
        autofit_type (str, optional): Specifies the type of auto-fit to apply. Supported values are:
                                      - "Shape": Adjusts the shape size to fit the text.
                                      - "Text": Shrinks/enlarges the text to fit the shape.
                                      - "None": No auto-fit behavior is applied.
                                      Defaults to "Shape".

    Returns:
        Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the auto-fit behavior was successfully applied, False otherwise.
            - message (str): Description of the result or error.
            - autofit_info (dict, optional): Additional information about the operation, such as:
                                             - slide_number (int)
                                             - shape_index (int)
                                             - applied_autofit_type (str)

    Raises:
        ShapeError: If the operation fails due to:
                      - Invalid file path or corrupted PowerPoint file
                      - Slide or shape index out of range
                      - Shape does not contain a text frame
                      - Unsupported auto-fit type provided
                      - Library-specific errors during auto-fit configuration
    """
    try:
        full_path = get_ppt_path(filepath)
        from .shape import set_autofittext as set_autofittext_impl
        result = set_autofittext_impl(
            filepath=full_path,
            slide_num=slide_num,
            shape_num = shape_num,
            autofit_type = autofit_type
        )
        return result
    except ShapeError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error:{e}")
        raise

@mcp.tool()
def set_verticaltext(
        filepath:str,
        slide_num:int = 0,
        shape_num:int = 0,
        verticaltext_type:str = "Vertical270"
) -> dict[str,Any]:
    """
    Sets the text orientation to vertical for the text inside a specified shape on a given slide in a PowerPoint presentation.

    This function loads the input presentation and changes the text direction of the specified shape's text frame 
    to a vertical layout based on the provided vertical text type.

    Parameters:
        filepath (str): Path to the input PowerPoint file (.pptx).
        slide_num (int, optional): Index of the slide containing the target shape. Defaults to 0.
        shape_num (int, optional): Index of the shape on the slide whose text orientation will be changed. Defaults to 0.
        verticaltext_type (str, optional): Specifies the vertical text orientation mode. Supported values are:
                                           - "Vertical": Text is rotated 90 degrees clockwise.
                                           - "Vertical270": Text is rotated 270 degrees clockwise (reads from top to bottom).
                                           Defaults to "Vertical270".

    Returns:
        Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the vertical text orientation was successfully applied, False otherwise.
            - message (str): Description of the result or error.
            - verticaltext_info (dict, optional): Additional information about the operation, such as:
                                                 - slide_number (int)
                                                 - shape_index (int)
                                                 - applied_verticaltext_type (str)

    Raises:
        ShapeError: If the operation fails due to:
                           - Invalid file path or corrupted PowerPoint file
                           - Slide or shape index out of range
                           - Shape does not contain a text frame
                           - Unsupported vertical text type provided
                           - Library-specific errors during text orientation change
    """
    try:
        full_path = get_ppt_path(filepath)
        from .shape import set_verticaltext as set_verticaltext_impl
        result = set_verticaltext_impl(
            filepath=full_path,
            slide_num=slide_num,
            shape_num = shape_num,
            verticaltext_type = verticaltext_type
        )
        return result
    except ShapeError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error:{e}")
        raise

@mcp.tool()
def set_text_color(
        filepath:str,
        slide_num:int = 0,
        shape_num:int = 0,
        color: str = None
) -> dict[str,Any]:
    """
    Sets the text color of a specified shape on a specified slide in a PowerPoint presentation.

    Parameters:
        filepath (str): Path to the PowerPoint (.pptx) file.
        slide_num (int, optional): Index of the slide containing the shape (starting from 0). Defaults to 0.
        shape_num (int, optional): Index of the shape within the slide (starting from 0). Defaults to 0.
        color (str, optional): Color to apply to the text (e.g., "red", "#FF5733"). If None, default color is used.

    Returns:
        dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): Whether the operation was successful.
            - message (str): Description of the result or error.
            - details (dict): Optional additional information (e.g., slide/shape index, applied color).

    Raises:
        ShapeError: If the slide, shape, or text frame cannot be accessed.
    """
    try:
        full_path = get_ppt_path(filepath)
        from .shape import set_text_color as set_text_color_impl
        result = set_text_color_impl(
            filepath=full_path,
            slide_num=slide_num,
            shape_num = shape_num,
            color = color
        )
        return result
    except ShapeError as e:
        return f"Error:{str(e)}"
    except Exception as e:
        logger.error(f"Error:{e}")
        raise


@mcp.tool()
def convert_pptx(
        filepath: str,
        output_filepath: str,
        format_type: str,  
) -> str:
    """
    Converts Ppt file to different formats.

    Supported formats:
    - pdf: Convert to PDF document
    - html: Convert to HTML document
    - image: Convert to image file png

    Parameters:
        filepath (str): Path to the Excel file
        format_type (str): Target format type (pdf, html, image)
        output_filepath (str): Path for the output file

    Returns:
        str: Success message or error description
    """
    try:
        full_path = get_ppt_path(filepath)
        output_path = get_ppt_path(output_filepath)
        
        result = convert_presentation_impl(
            filepath=full_path,
            output_filepath=output_path,
            format_type=format_type
        )
        
        return result
    except ConversionError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error converting file: {e}")
        raise ConversionError(f"Failed to convert Ppt file: {str(e)}")
    
async def run_server():
    """Run the Spire.Ppt MCP Server."""
    try:
        logger.info(f"Starting Spire.Ppt MCP Server (files directory: {PPT_FILES_PATH})")
        await mcp.run_sse_async()
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
        await mcp.shutdown()
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Server shutdown complete")
