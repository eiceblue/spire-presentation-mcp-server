import logging
from typing import Any,Dict

from spire.presentation import *

from .exceptions import ConversionError

logger = logging.getLogger(__name__)

def convert_presentation(
        filepath:str,
        output_filepath:str,
        format_type:str,
        options:Dict[str,Any] = None,
) -> dict[str,Any]:
    """
    Convert Ppt presentation different formats.
        
    Args:
    filepath: Source Ppt file path
    output_filepath: Target output file path
    format_type: Target format (pdf,, html, image, txt, pptx，etc.)
    options: Format-specific options
            
    Returns:
    Dictionary with operation status
    """
    try:
        #Create an instance of presentation document
        ppt = Presentation()

        #Load file
        ppt.LoadFromFile(filepath)

        # Ensure output directory exists
        output_dir = os.path.dirname(output_filepath)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        # Handle format-specific conversion
        format_type = format_type.lower()

        if format_type == 'pdf':
            ppt.SaveToFile(output_filepath,FileFormat.PDF)

        elif format_type == 'html':
            ppt.SaveToFile(output_filepath,FileFormat.Html)

        elif format_type == 'ofd':
            ppt.SaveToFile(output_filepath,FileFormat.OFD)

        elif format_type == 'xps':
            ppt.SaveToFile(output_filepath,FileFormat.XPS)

        elif format_type == 'svg':
            for index,slide in enumerate(ppt.Slides):
                fileName =  "ToSVG-"+str(index)+".svg"
                svgStream = slide.SaveToSVG()
                svgStream.Save(fileName)

        elif format_type == 'image':
            #Save PPT document to images
            for i, slide in enumerate(ppt.Slides):
                fileName ="ToImage_img_"+str(i)+".png"
                image = slide.SaveAsImage()
                image.Save(fileName)
                image.Dispose()

        return {
            "message": f"Ppt file successfully converted to {format_type.upper()}: {output_filepath}",
            "source_file": filepath,
            "output_file": output_filepath,
            "format": format_type
        }

    except ConversionError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to convert Ppt file: {e}")
        raise ConversionError(f"Failed to convert Ppt file: {str(e)}")