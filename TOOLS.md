# Spire.Presentaion MCP Server Tools

This document provides detailed information about all available tools in the Spire.Presntation MCP Server.

## Presentaion Operations

### create_presentation

Creates a new Ppt presentation.

```python
create_presentation(filepath:str) -> str:
```

- `filepath`: Path where the new presentation will be saved
- Returns: Success message with the created presentation path

## Slide Operations

### create_slide

Creates a new slide in an existing presentaion.

```python
create_slide(filepath:str) -> str:
```

- `filepath`: Path to the Ppt file
- Returns: Success message confirming slide creation

### delete_slide

Deletes a slide from an existing presentation.

```python
delete_slide(filepath:str,slide_num:int) -> dict[str,Any]:
```

- `filepath`: Path to the PPT file
- `slide_num (int)`: Index of the slide to delete (starting from 0)
- Returns: Success message confirming slide deletion

### change_slide_position

Changes the position of a specified slide in the PowerPoint presentation to a new index.

```python
def change_slide_position(
    filepath: str,
    slide_num:int,
    slide_number:int
) -> dict[str,Any]:
```

- `filepath (str)`: Path to the input PowerPoint file (.pptx).
- `slide_num (int)`: Index of the slide to be moved (current position, 0-based or library-dependent).
- `slide_number (int)`: Target index where the slide should be placed (0-based or library-dependent).
- Returns: Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the slide was successfully moved, False otherwise.
            - message (str): Description of the result or error.
            - move_info (dict, optional): Additional information about the slide movement, such as:
                                          - original_position (int)
                                          - new_position (int)

### add_image_in_master

Adds an image to a specified slide master in a PowerPoint presentation.

```python
def add_image_in_master(
        filepath: str, 
        image_filepath: str,
        master_num:int = 0,
        x:float = 0,
        y:float = 0,
        width:float = 200,
        height:float = 200
) -> dict[str,Any]:
```

- `filepath (str)`: Path to the input PowerPoint file (.pptx).
- `image_filepath (str)`: Path to the image file to be inserted into the slide master.
                              Supported formats typically include: .png, .jpg, .gif, etc.
- `master_num (int, optional)`: Index of the slide master to which the image will be added.
                                    Defaults to 0 (the first master).
- `x (float, optional)`: X-coordinate (in points) for the top-left corner of the image.
                             Defaults to 0.
- `y (float, optional)`: Y-coordinate (in points) for the top-left corner of the image.
                             Defaults to 0.
- `width (float, optional)`: Width of the inserted image (in points). Defaults to 200.
- `height (float, optional)`: Height of the inserted image (in points). Defaults to 200.
- Returns: Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the image was successfully added to the master, False otherwise.
            - message (str): Description of the result or error.
            - master_info (dict, optional): Additional information about the operation, such as:
                                            - master_index (int)
                                            - image_position (tuple[float, float]): (x, y)
                                            - image_size (tuple[float, float]): (width, height)

## Shape Operations

### add_shape

Adds a geometric shape to a specified slide in a PowerPoint presentation.

```python
add_shape(
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
```

- `filepath (str)`: Path to the PowerPoint file (.pptx).
- `slide_num (int, optional)`: Slide number where the shape will be added.
                                   Defaults to 0 (typically the first slide).
- `x (float, optional)`: X-coordinate of the shape's top-left corner in pixels. Defaults to 0.
- `y (float, optional)`: Y-coordinate of the shape's top-left corner in pixels. Defaults to 0.
- `width (float, optional)`: Width of the shape in pixels. Defaults to 200.
- `height (float, optional)`: Height of the shape in pixels. Defaults to 200.
- `shape_type (str, optional)`: Type of shape to add. Supported types depend on the underlying library.Defaults to 'Rectangle'.
- `line_color (str, optional)`: Outline (stroke) color of the shape in hex format (e.g., '#FF5733').
                                    If not provided, the default outline style is used or no outline may be applied.
- `fill_color (str, optional)`: Fill color of the shape in hex format (e.g., '#C0C0C0').
                                    If not provided, the shape will have no fill.
- Returns: Dict[str, Any]: A dictionary containing operation result with the following keys:
            - success (bool): Whether the operation was successful.
            - message (str): Detailed message or error description.
            - shape_info (dict, optional): Information about the added shape, such as:
                                           - type (str)
                                           - position (x, y)
                                           - size (width, height)
                                           - colors (line and fill if applicable)

### delete_shape

Deletes a shape from a specified slide in a PowerPoint presentation.

```python
def delete_shape(filepath:str,slide_num:int,shape_num:int) -> dict[str,Any]:
```

- `filepath (str)`: Path to the PowerPoint file.
- `slide_num (int)`: Slide number from which the shape will be deleted.
- `shape_num (int)`: The index or identifier of the shape to delete. 
- Returns: Dict[str, Any]: A dictionary containing operation result with the following keys:
            success (bool): Whether the operation was successful.
            message (str): Detailed message or error description.
            shape_info (dict, optional): Information about the deleted shape (if available and applicable).

### add_text_shape

Adds a new text shape or updates an existing one on a specified slide in a PowerPoint presentation.

```python
add_text_shape(filepath:str,slide_num:int,shape_num:int = None,text:str = "") -> dict[str,Any]:
```

- `filepath (str)`: Path to the PowerPoint file.
- `slide_num (int)`: Slide number where the text shape will be added or updated.
- `shape_num (int)`: Index or identifier of an existing shape to update.
                    If None, a new text shape will be created. Defaults to None.
-`text (str, optional)`: Text content to insert into the shape. Defaults to an empty string.
- Returns: Dict[str, Any]: A dictionary containing operation result with the following keys:
            - success (bool): Whether the operation was successful.
            - message (str): Detailed message or error description.
            - shape_info (dict, optional): Information about the added or updated text shape,
                                          such as its index, position, size, and text content.

### shape_to_image

Converts shapes on a specified slide of a PowerPoint presentation into an image file.

```python
add_text_shape(filepath:str,slide_num:int,shape_num:int = None,text:str = "") -> dict[str,Any]:
```

- `filepath (str)`: Path to the input PowerPoint file (.pptx).
- `slide_num (int)`: Slide number to convert into an image.
                         The index typically starts at 0 or 1 depending on the library or backend used.
- `output_filepath (str)`: Full path where the output image will be saved (e.g., 'output/slide.png').
                               The file extension determines the image format (e.g., .png).
- Returns: Dict[str, Any]: A dictionary containing operation result with the following keys:
            - success (bool): Whether the operation was successful.
            - message (str): Detailed message or error description.
            - image_info (dict, optional): Additional information about the generated image, such as:
                                           - file_path (str): Path to the saved image.
                                           - dimensions (tuple): Width and height of the image.
                                           - slide_number (int): Index of the exported slide.

### set_shape_fill_picture

Sets the fill of a specified shape on a slide with an image from the given URL or local path.

```python
def set_shape_fill_picture(
    filepath:str,
    slide_num:int,
    shape_num:int,
    picture_url:str
) -> dict[str,Any]:
```

- `filepath (str)`: Path to the input PowerPoint file (.pptx).
- `slide_num (int)`: Index of the slide containing the target shape.
                         (Index typically starts at 0 depending on the library used)
- `shape_num (int)`: Index of the shape within the slide to apply the image fill.
- `picture_url (str)`: Path or URL to the image file. Supported formats typically include:
                           .png, .jpg, .gif, .bmp, etc., depending on backend support.
- Returns: Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the image was successfully applied, False otherwise.
            - message (str): Description of the result or error.
            - shape_info (dict, optional): Additional information about the updated shape, such as:
                                           - slide_number (int)
                                           - shape_index (int)
                                           - image_source (str): The original image path/URL used

### get_shape_titles

Extracts titles or text content from shapes in all slides of a PowerPoint presentation,
    and saves the extracted text into a plain text (.txt) file.

```python
def get_shape_titles(
        filepath:str,
        output_filepath:str
) -> dict[str,Any]:
```

- `filepath (str)`: Path to the source PowerPoint file (.pptx).
- `output_filepath (str)`: Path where the extracted shape texts will be saved as a .txt file.
                               The file extension should typically be .txt.
- Returns: Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the image was successfully applied, False otherwise.
            - message (str): Description of the result or error.
            - shape_info (dict, optional): Additional information about the updated shape, such as:
                                           - slide_number (int)
                                           - shape_index (int)
                                           - image_source (str): The original image path/URL used

### group_shapes

Groups a list of shape objects on a specified slide into a single shape group within the PowerPoint presentation.

```python
def group_shapes(
        filepath:str,
        slide_num:int,
        shape_num_list:List[int] = []
) -> dict[str,Any]:
```

- `filepath (str)`: Path to the input PowerPoint file (.pptx).
- `slide_num (int)`: Index of the slide containing the shapes to be grouped.
                         (Index typically starts at 0 depending on the library used)
- `shape_num_list (List[int], optional)`: A list of indices of shapes on the slide to be grouped. Defaults to an empty list.
- Returns: Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the shapes were successfully grouped, False otherwise.
            - message (str): Description of the result or error.
            - group_info (dict, optional): Additional information about the grouping, such as:
                                           - slide_number (int)
                                           - grouped_shape_count (int): Number of shapes grouped

### ungroup_shapes

Ungroups a grouped shape on a specified slide in a PowerPoint presentation.

```python
def ungroup_shapes(
        filepath:str,
        slide_num:int,
        shape_num:int
) -> dict[str,Any]:
```

- `filepath (str)`: Path to the input PowerPoint file (.pptx).
- `slide_num (int)`: Index of the slide containing the group shape.
                         (Index typically starts at 0 depending on the library used)
- `shape_num (int)`: Index of the shape on the slide that should be ungrouped.
- Returns: Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the shape was successfully ungrouped, False otherwise.
            - message (str): Description of the result or error.
            - ungroup_info (dict, optional): Additional information about the operation, such as:
                                             - slide_number (int)
                                             - ungrouped_shape_index (int)
                                             - number_of_ungrouped_shapes (int): Number of individual shapes released after ungrouping

### set_alignment

Sets the text alignment for a specific paragraph within a shape on a given slide in a PowerPoint presentation.

```python
def set_alignment(
        filepath:str,
        slide_num:int = 0,
        shape_num:int = 0,
        paragraph_num:int = 0,
        text_alignment_type:str = "Left"
) -> dict[str,Any]:
```

- `filepath (str)`: Path to the input PowerPoint file (.pptx).
- `slide_num (int, optional)`: Index of the slide containing the target shape. Defaults to 0.
- `shape_num (int, optional)`: Index of the shape within the slide that contains the target paragraph. Defaults to 0.
- `paragraph_num (int, optional)`: Index of the paragraph within the shape whose alignment needs to be changed. Defaults to 0.
- `text_alignment_type (str, optional)`: Desired text alignment type. Possible values are 'Left', 'Right', 'Center', 'Justify'.
                                             Defaults to 'Left'.
- Returns: Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the text alignment was successfully applied, False otherwise.
            - message (str): Description of the result or error.
            - alignment_info (dict, optional): Additional information about the operation, such as:
                                               - slide_number (int)
                                               - shape_index (int)
                                               - paragraph_index (int)
                                               - applied_alignment (str)

### append_html

Appends HTML-formatted text content into a specified text frame within a shape on a given slide.

```python
def append_html(
        filepath:str,
        slide_num:int = 0,
        shape_num:int = 0,
        code_html:str = " "
) -> dict[str,Any]:
```

- `filepath (str)`: Path to the input PowerPoint file (.pptx).
- `slide_num (int, optional)`: Index of the slide containing the target shape. Defaults to 0.
- `shape_num (int, optional)`: Index of the shape on the slide that contains the text frame. Defaults to 0.
- `code_html (str, optional)`: HTML-formatted string to be appended into the text frame. 
                                   Defaults to a single space (" ").
- Returns: Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the HTML content was successfully appended, False otherwise.
            - message (str): Description of the result or error.
            - html_info (dict, optional): Additional information about the operation, such as:
                                           - slide_number (int)
                                           - shape_index (int)
                                           - applied_html (str)

### set_autofittext

Sets the auto-fit behavior for the text within a specified shape on a given slide in a PowerPoint presentation.

```python
def set_autofittext(
        filepath:str,
        slide_num:int = 0,
        shape_num:int = 0,
        autofit_type:str = "Shape"
) -> dict[str,Any]:
```

- `filepath (str)`: Path to the input PowerPoint file (.pptx).
- `slide_num (int, optional)`: Index of the slide containing the target shape. Defaults to 0.
- `shape_num (int, optional)`: Index of the shape on the slide whose text auto-fit behavior will be set. Defaults to 0.
- `autofit_type (str, optional)`: Specifies the type of auto-fit to apply. Defaults to "Shape".
- Returns: Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the auto-fit behavior was successfully applied, False otherwise.
            - message (str): Description of the result or error.
            - autofit_info (dict, optional): Additional information about the operation, such as:
                                             - slide_number (int)
                                             - shape_index (int)
                                             - applied_autofit_type (str)

### set_verticaltext

Sets the text orientation to vertical for the text inside a specified shape on a given slide in a PowerPoint presentation.

```python
def set_verticaltext(
        filepath:str,
        slide_num:int = 0,
        shape_num:int = 0,
        verticaltext_type:str = "Vertical270"
) -> dict[str,Any]:
```

- `filepath (str)`: Path to the input PowerPoint file (.pptx).
- `slide_num (int, optional)`: Index of the slide containing the target shape. Defaults to 0.
- `shape_num (int, optional)`: Index of the shape on the slide whose text orientation will be changed. Defaults to 0.
- `verticaltext_type (str, optional)`: Specifies the vertical text orientation mode. Defaults to "Vertical270".
- Returns: Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the vertical text orientation was successfully applied, False otherwise.
            - message (str): Description of the result or error.
            - verticaltext_info (dict, optional): Additional information about the operation, such as:
                                                 - slide_number (int)
                                                 - shape_index (int)
                                                 - applied_verticaltext_type (str)


### set_text_color

Sets the text color of a specified shape on a specified slide in a PowerPoint presentation.

```python
def set_text_color(
        filepath:str,
        slide_num:int = 0,
        shape_num:int = 0,
        color: str = None
) -> dict[str,Any]:
```

- `filepath (str)`: Path to the PowerPoint (.pptx) file.
- `slide_num (int, optional)`: Index of the slide containing the shape (starting from 0). Defaults to 0.
- `shape_num (int, optional)`: Index of the shape within the slide (starting from 0). Defaults to 0.
- `color (str, optional)`: Color to apply to the text (e.g., "red", "#FF5733"). If None, default color is used.
- Returns: dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): Whether the operation was successful.
            - message (str): Description of the result or error.
            - details (dict): Optional additional information (e.g., slide/shape index, applied color).

## Chart Operations

### add_chart

Adds a new chart to a specified location on a slide in a PowerPoint presentation.

```python
def add_chart(
    filepath:str,
    slide_num:int,
    x:float = 0,
    y:float = 0,
    width:float = 200,
    height:float = 200,
    chart_type:str = "Pie"
) -> dict[str,Any]:
```

- `filepath (str)`: Path to the PowerPoint file (.pptx).
- `slide_num (int)`: Slide number where the chart will be added.
                         The index typically starts at 0 or 1 depending on the library used.
- `x (float, optional)`: X-coordinate of the chart's top-left corner in pixels. Defaults to 0.
- `y (float, optional)`: Y-coordinate of the chart's top-left corner in pixels. Defaults to 0.
- `width (float, optional)`: Width of the chart in pixels. Defaults to 200.
- `height (float, optional)`: Height of the chart in pixels. Defaults to 200.
- `chart_type (str, optional)`: Type of chart to add. Supported types depend on the underlying library.
                                    Common options include: 'Bar', 'Column', 'Pie', 'Line', 'Scatter', etc.
                                    Defaults to 'Pie'.
- Returns: Dict[str, Any]: A dictionary containing operation result with the following keys:
            - success (bool): Whether the operation was successful.
            - message (str): Detailed message or error description.
            - chart_info (dict, optional): Information about the added chart, such as:
                                           - chart_type (str)
                                           - position (x, y)
                                           - size (width, height)
                                           - reference ID or other metadata (if applicable)

## SmartArt Operations

### create_smartart

Adds a SmartArt graphic to a specified location on a slide in a PowerPoint presentation.

```python
def create_smartart(
    filepath:str,
    slide_num:int,
    x:float = 0,
    y:float = 0,
    width:float = 200,
    height:float = 200,
    layout_type:str = "Gear"
) -> dict[str,Any]:
```

- `filepath (str)`: Path to the PowerPoint file (.pptx).
- `slide_num (int)`: Slide number where the SmartArt will be added.
                         The index typically starts at 0 or 1 depending on the library used.
- `x (float, optional)`: X-coordinate of the SmartArt's top-left corner in pixels. Defaults to 0.
- `y (float, optional)`: Y-coordinate of the SmartArt's top-left corner in pixels. Defaults to 0.
- `width (float, optional)`: Width of the SmartArt graphic in pixels. Defaults to 200.
- `height (float, optional)`: Height of the SmartArt graphic in pixels. Defaults to 200.
- `layout_type (str, optional)`: Type of SmartArt layout to add. Supported types depend on the underlying library.
                                     Common examples include: 'Gear', 'Bubbles', 'List', 'Process', 'Hierarchy', etc.
                                     Defaults to 'Gear'.
- Returns: Dict[str, Any]: A dictionary containing operation result with the following keys:
            - success (bool): Whether the operation was successful.
            - message (str): Detailed message or error description.
            - smartart_info (dict, optional): Information about the added SmartArt graphic, such as:
                                              - layout_type (str)
                                              - position (x, y)
                                              - size (width, height)
                                              - reference ID or other metadata (if applicable)

## Table Operations

### create_table

Creates a table on a specified slide in a PowerPoint presentation.

```python
def create_table(
    filepath:str,
    slide_num:int,
    x:float = 0,
    y:float = 0,
    widths:List[float] = [50,50],
    heights:List[float] = [20,20]
) -> dict[str,Any]:
```

- `filepath (str)`: Path to the PowerPoint file (.pptx).
- `slide_num (int)`: Index of the slide where the table will be added.
                         The index typically starts at 0 or 1 depending on the library used.
- `x (float, optional)`: X-coordinate of the table's top-left corner in pixels. Defaults to 0.
- `y (float, optional)`: Y-coordinate of the table's top-left corner in pixels. Defaults to 0.
- `widths (List[float], optional)`: A list specifying the width of each column in pixels. 
                                        Defaults to [50, 50] (two columns, 50 pixels wide each).
- `heights (List[float], optional)`: A list specifying the height of each row in pixels.
                                         Defaults to [20, 20] (two rows, 20 pixels high each).
- Returns: Dict[str, Any]: A dictionary containing the result of the operation with the following keys:
            - success (bool): Whether the table was successfully created.
            - message (str): A detailed message or error description.
            - table_info (dict, optional): Information about the created table, such as:
                                           - position (x, y)
                                           - dimensions (number of rows and columns)
                                           - total_width (float): Total width of the table.
                                           - total_height (float): Total height of the table.
                                           - column_widths (List[float])
                                           - row_heights (List[float])

### add_text_table

Adds or updates text content in a table shape on a specified slide of a PowerPoint presentation.

```python
def add_text_table(
        filepath:str,
        slide_num:int,
        shape_num:int,
        data_str:List[str] = ["", "", "", ""]
) -> dict[str,Any]:
```

- `filepath (str)`: Path to the PowerPoint file (.pptx).
- `slide_num (int)`: Index of the slide containing the table (0-based index unless otherwise defined by the library).
- `shape_num (int)`: Index of the shape on the slide; expected to be a table shape.
- `data_str (List[str], optional)`: A list of strings used to fill the table cells row by row.
                                          Defaults to ["", "", "", ""].
- Returns: Dict[str, Any]: A dictionary containing the result of the operation:
            - success (bool): True if the operation was successful, False otherwise.
            - message (str): Description of the result or error.
            - table_info (dict, optional): Additional information about the updated table, such as:
                                           - rows (int): Number of rows in the table.
                                           - columns (int): Number of columns in the table.
                                           - filled_data (List[List[str]]): The 2D list used to fill the table.

## Conversion Operations

### convert_pptx

Converts Ppt file to different formats.

```python
def convert_pptx(
        filepath: str,
        output_filepath: str,
        format_type: str,  
) -> str:
```
Supported formats:
    - pdf: Convert to PDF document
    - html: Convert to HTML document
    - image: Convert to image file png

- `filepath (str)`: Path to the Excel file
- `format_type (str)`: Target format type (pdf, html, image)
- `output_filepath (str)`: Path for the output file
- Returns: Success message or error description
