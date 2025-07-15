import logging
from typing import Any

from spire.presentation import *

from .exceptions import SlideError

logger = logging.getLogger(__name__)

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
    Add pictures to master.
    """
    try:
        ppt = Presentation()
        ppt.LoadFromFile(filepath)

        #Get the master collection
        master = ppt.Masters[master_num]

        #Append image to slide master
        image = image_filepath
        rff = RectangleF.FromLTRB (x, y, width + x, height + y)
        pic = master.Shapes.AppendEmbedImageByPath(ShapeType.Rectangle, image, rff)
        pic.Line.FillFormat.FillType = FillFormatType.none

        #Add new slide to presentation
        ppt.Slides.Append()

        # 保存更改
        ppt.SaveToFile(filepath,FileFormat.Pptx2019)

        return {"message": f"master add pictures successfully"}

    except SlideError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Adding image failed: {e}")
        raise SlideError(str(e))
    
def append_slide_with_master_layout(filepath: str) -> dict:
    try:
        #Create a PPT document
        presentation = Presentation()
        #Load the document from disk
        presentation.LoadFromFile(filepath)
        #Get the master
        master = presentation.Masters[0]
        #Get master layout slides
        masterLayouts = master.Layouts
        layoutSlide = masterLayouts[1]
        #Append a rectangle to the layout slide
        shape = layoutSlide.Shapes.AppendShape(ShapeType.Rectangle, RectangleF.FromLTRB (10, 50, 110, 130))
        #Add a text into the shape and set the style
        shape.Fill.FillType = FillFormatType.none
        shape.AppendTextFrame("Layout slide 1")
        shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = TextFont("Arial Black")
        shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = FillFormatType.Solid
        shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.get_CadetBlue()
        #Append new slide with master layout
        presentation.Slides.Append(presentation.Slides[0], master.Layouts[1])
        #Another way to append new slide with master layout
        presentation.Slides.Insert(2, presentation.Slides[1], master.Layouts[1])
        #Save the document
        presentation.SaveToFile(filepath)
        return {"message": f"append successfully"}

    except SlideError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Add failed: {e}")
        raise SlideError(str(e))
    
def apply_slide_master(filepath: str, image_filepath: str) -> dict[str,Any]:
    try:
        #Create an instance of presentation document
        ppt = Presentation()
        #Load file
        ppt.LoadFromFile(filepath)
        #Get the first slide master from the presentation
        masterSlide = ppt.Masters[0]
        #Customize the background of the slide master
        backgroundPic = image_filepath
        rect = RectangleF.FromLTRB (0, 0, ppt.SlideSize.Size.Width, ppt.SlideSize.Size.Height)
        masterSlide.SlideBackground.Fill.FillType = FillFormatType.Picture
        image = masterSlide.Shapes.AppendEmbedImageByPath (ShapeType.Rectangle, backgroundPic, rect)
        masterSlide.SlideBackground.Fill.PictureFill.Picture.EmbedImage = image.PictureFill.Picture.EmbedImage
        #Change the color scheme
        masterSlide.Theme.ColorScheme.Accent1.Color = Color.get_Red()
        masterSlide.Theme.ColorScheme.Accent2.Color = Color.get_RosyBrown()
        masterSlide.Theme.ColorScheme.Accent3.Color = Color.get_Ivory()
        masterSlide.Theme.ColorScheme.Accent4.Color = Color.get_Lavender()
        masterSlide.Theme.ColorScheme.Accent5.Color = Color.get_Black()
        #Save the document
        ppt.SaveToFile(filepath)
        return {"message": f"apply successfully"}

    except SlideError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"apply failed: {e}")
        raise SlideError(str(e))
    
def change_slide_position(filepath: str,slide_num:int,slide_number:int) -> dict[str,Any]:
    try:
        #Create a PPT document
        presentation = Presentation()
        #Load the document from disk
        presentation.LoadFromFile(filepath)
        
        slide = presentation.Slides[slide_num]
        slide.SlideNumber = slide_number
        #Save the document
        presentation.SaveToFile(filepath,FileFormat.Pptx2019)
        return {"message": f"change successfully"}

    except SlideError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"change failed: {e}")
        raise SlideError(str(e))
    
def append_slide(filepatth:str) -> dict[str,Any]:
    try:
        ppt = Presentation()
        ppt.LoadFromFile(filepatth)

        ppt.Slides.Append()

        ppt.SaveToFile(filepatth,FileFormat.Pptx2019)
        return {"message": "append successfully",
                "slide":"some slide information"}

    except SlideError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"append failed: {e}")
        raise SlideError(str(e))
    
def delete_slide(filepatth:str,slide_num:int) -> dict[str,Any]:
    try:
        ppt = Presentation()
        ppt.LoadFromFile(filepatth)

        if slide_num > ppt.Slides.Count:
            raise SlideError(f"length {slide_num} greater than slide count")
        
        ppt.Slides.RemoveAt(slide_num)

        ppt.SaveToFile(filepatth,FileFormat.Pptx2019)
        return {"message": f"delete {slide_num} slide successfully"}
    except SlideError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"delete failed: {e}")
        raise SlideError(str(e))