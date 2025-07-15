import logging
import os
from typing import Any

from spire.presentation import *

from .exceptions import ShapeError

logger = logging.getLogger(__name__)

def add_line_to_slide(filepath:str) -> dict:
    try:
        #Create a PPT document
        presentation = Presentation()
        presentation.LoadFromFile(filepath)
        #Get the first slide
        slide = presentation.Slides[0]
        #Add a line in the slide
        line = slide.Shapes.AppendShape(ShapeType.Line, RectangleF.FromLTRB (50, 100, 350, 100))
        #Set color of the line
        line.ShapeStyle.LineColor.Color = Color.get_Red()
        #Save the document
        presentation.SaveToFile(filepath,FileFormat.Pptx2019)
        return {"message": f"add successfully"}

    except ShapeError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"add failed: {e}")
        raise ShapeError(str(e))
    
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
    try:
        
        #Create a PPT document
        ppt = Presentation()
        ppt.LoadFromFile(filepath)
        
        slide = ppt.Slides[slide_num]
        
        rect = RectangleF.FromLTRB (x, y, width + x, height + y)
        
        type1 = None
        for type in ShapeType:
            if type.name == shape_type:
                type1 = type
        
        if type1 == None:
            type1 = ShapeType.Rectangle

        
        shape = slide.Shapes.AppendShape(type1,rect)
        
        if fill_color != None:
            shape.Fill.FillType = FillFormatType.Solid
            if fill_color.startswith('#'):
                fill_color = fill_color[1:]
            if len(fill_color) == 6:
                r = int(fill_color[0:2], 16)
                g = int(fill_color[2:4], 16)
                b = int(fill_color[4:6], 16)
                shape.Fill.SolidColor.Color = Color.FromRgb(r, g, b)
        
        if line_color != None:
            shape.Line.FillType = FillFormatType.Solid
            if line_color.startswith('#'):
                line_color = line_color[1:]
            if len(line_color) == 6:
                r = int(line_color[0:2], 16)
                g = int(line_color[2:4], 16)
                b = int(line_color[4:6], 16)
                shape.Line.SolidFillColor.Color = Color.FromRgb(r, g, b)
        
        #Save the document
        ppt.SaveToFile(filepath,FileFormat.Pptx2019)
        return {"message": f"add successfully"}

    except ShapeError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"add failed: {e}")
        raise ShapeError(str(e))

def delete_shape(filepath:str,slide_num:int,shape_num:int) -> dict[str,Any]:
    try:
        #Create a PPT document
        ppt = Presentation()
        ppt.LoadFromFile(filepath)
            
        slide = ppt.Slides[slide_num]

        if shape_num > ppt.Slides[slide_num].Shapes.Count:
            raise ShapeError(f"length {shape_num} greater than shape count")
        
        slide.Shapes.RemoveAt(shape_num)

        ppt.SaveToFile(filepath,FileFormat.Pptx2019)
        return {"message": f"delete successfully"}

    except ShapeError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"delete failed: {e}")
        raise ShapeError(str(e))
    
def add_text_shape(filepath:str,slide_num:int,shape_num:int = None,text:str = "") -> dict[str,Any]:
    try:
        #Create a PPT document
        ppt = Presentation()
        ppt.LoadFromFile(filepath)
            
        slide = ppt.Slides[slide_num]
        
        shape = None
        if shape_num == None:
            shape = slide.Shapes.AppendShape(ShapeType.Rectangle,RectangleF.FromLTRB (0, 0, 200, 200))
            shape.Fill.FillType = FillFormatType.none
            shape.Line.FillType = FillFormatType.none
        else:
            if shape_num > ppt.Slides[slide_num].Shapes.Count:
                raise ShapeError(f"length {shape_num} greater than shape count")
            shape = slide.Shapes[shape_num]

        shape.TextFrame.Text = text
        ppt.SaveToFile(filepath,FileFormat.Pptx2019)
        return {"message": f"add text successfully"}

    except ShapeError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"add text failed: {e}")
        raise ShapeError(str(e))
    
def shape_to_image(filepath:str,slide_num:int,output_filepath:str) -> dict[str,Any]:
    try:
        #Create a PPT document
        ppt = Presentation()
        ppt.LoadFromFile(filepath)
            
        slide = ppt.Slides[slide_num]

        new_path = output_filepath.rsplit('.', 1)[0]
        
        if not os.path.exists(new_path):
            os.mkdir(new_path)

        for i, unusedItem in enumerate(slide.Shapes):
            fileName = new_path + "//" + "ShapeToImage-"+str(i)+".png"
            #Save shapes as images
            image = slide.Shapes.SaveAsImage(i)
            image.Save(fileName)
            image.Dispose()

        return {"message": f"successfully"}

    except ShapeError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"failed: {e}")
        raise ShapeError(str(e))
    
def fill_shape_with_picture(filepath:str,slide_num:int,shape_num:int,picture_url:str) -> dict[str,Any]:
    try:
        #Create a PPT document
        ppt = Presentation()
        ppt.LoadFromFile(filepath)

        if slide_num > ppt.Slides.Count:
                raise ShapeError(f"length {slide_num} greater than slide count")

        slide = ppt.Slides[slide_num]

        if shape_num > ppt.Slides[slide_num].Shapes.Count:
                raise ShapeError(f"length {shape_num} greater than shape count")
        
        shape = slide.Shapes[shape_num]

        shape.Fill.FillType = FillFormatType.Picture
        shape.Fill.PictureFill.Picture.Url = picture_url
        shape.Fill.PictureFill.FillType = PictureFillType.Stretch

        ppt.SaveToFile(filepath,FileFormat.Pptx2019)
        return {"message": f"add successfully"}

    except ShapeError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"add failed: {e}")
        raise ShapeError(str(e))
    
def get_shape_titles(
        filepath:str,
        output_filepath:str
) -> dict[str,Any]:
    try:
         #Create a PPT document
        ppt = Presentation()
        ppt.LoadFromFile(filepath)

        #Instantiate a list of IShape objects
        shapelist = []
        #Loop through all sildes and all shapes on each slide
        for slide in ppt.Slides:
            for shape in slide.Shapes:
                if not isinstance(shape,ISmartArt):
                    if shape.Placeholder is not None:
                        #Get all titles
                        if shape.Placeholder.Type == PlaceholderType.Title:
                            shapelist.append(shape)
                        elif shape.Placeholder.Type == PlaceholderType.CenteredTitle:
                            shapelist.append(shape)
                        elif shape.Placeholder.Type == PlaceholderType.Subtitle:
                            shapelist.append(shape)
        #Loop through the list and get the inner text of all shapes in the list
        sb = []
        sb.append("Below are all the obtained titles:")
        for i, unusedItem in enumerate(shapelist):
            shape1 = shapelist[i] if isinstance(shapelist[i], IAutoShape) else None
            sb.append (shape1.TextFrame.Text)
        #Save to the Text file
        fp = open(output_filepath,"w")
        for s in sb:
            fp.write(s + "\n")
        fp.close()
        return {"message": f"add successfully"}

    except ShapeError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"add failed: {e}")
        raise ShapeError(str(e))
    
def group_shapes(
        filepath:str,
        slide_num:int,
        shape_num_list:List[int] = []
) -> dict[str,Any]:
    try:
        #Create a PPT document
        ppt = Presentation()
        ppt.LoadFromFile(filepath)

        slide = ppt.Slides[slide_num]
        
        if len(shape_num_list) < 1:
            raise ShapeError("Count less than or equal to one")
        
        shape_list = []
        for i in shape_num_list:
            shape = slide.Shapes[i]
            shape_list.append(shape)

        slide.GroupShapes(shape_list)
        ppt.SaveToFile(filepath,FileFormat.Pptx2019)
        return {"message": f"add successfully"}

    except ShapeError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"add failed: {e}")
        raise ShapeError(str(e))
    
def ungroup_shapes(
        filepath:str,
        slide_num:int,
        shape_num:int
) -> dict[str,Any]:
    try:
        #Create a PPT document
        ppt = Presentation()
        ppt.LoadFromFile(filepath)

        slide = ppt.Slides[slide_num]

        shape = slide.Shapes[shape_num]

        if isinstance(shape,GroupShape):
            slide.Ungroup(shape)
        else:
            raise ShapeError("Shape does not belong to groupshape")
        
        ppt.SaveToFile(filepath,FileFormat.Pptx2019)
        return {"message": f"successfully"}

    except ShapeError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"failed: {e}")
        raise ShapeError(str(e))
    
def set_alignment(
        filepath:str,
        slide_num:int = 0,
        shape_num:int = 0,
        paragraph_num:int = 0,
        text_alignment_type:str = "Left"
) -> dict[str,Any]:
    try:
        #Create a PPT document
        ppt = Presentation()
        ppt.LoadFromFile(filepath)

        slide = ppt.Slides[slide_num]

        shape = slide.Shapes[shape_num]

        type1 = None
        for type in TextAlignmentType:
            if type.name == text_alignment_type:
                type1 = type

        if type1 == None:
            type1 = TextAlignmentType.Left

        shape.TextFrame.Paragraphs[paragraph_num].Alignment = type1

        ppt.SaveToFile(filepath,FileFormat.Pptx2019)
        return {"message": f"successfully"}

    except ShapeError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"failed: {e}")
        raise ShapeError(str(e))

def append_html(
        filepath:str,
        slide_num:int = 0,
        shape_num:int = 0,
        code_html:str = " "
) -> dict[str,Any]:
    try:
        #Create a PPT document
        ppt = Presentation()
        ppt.LoadFromFile(filepath)

        slide = ppt.Slides[slide_num]

        shape = slide.Shapes[shape_num]

        #Clear default paragraphs 
        shape.TextFrame.Paragraphs.Clear()

        shape.TextFrame.Paragraphs.AddFromHtml(code_html)

        ppt.SaveToFile(filepath,FileFormat.Pptx2019)
        return {"message": f"successfully"}

    except ShapeError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"failed: {e}")
        raise ShapeError(str(e))

def set_autofittext(
        filepath:str,
        slide_num:int = 0,
        shape_num:int = 0,
        autofit_type:str = "Shape"
) -> dict[str,Any]:
    try:
        #Create a PPT document
        ppt = Presentation()
        ppt.LoadFromFile(filepath)

        slide = ppt.Slides[slide_num]

        shape = slide.Shapes[shape_num]

        type1 = None
        for type in TextAutofitType:
            if type.name == autofit_type:
                type1 = type

        if type1 == None:
            type1 = TextAutofitType.Shape

        shape.TextFrame.AutofitType = type1
        ppt.SaveToFile(filepath,FileFormat.Pptx2019)
        return {"message": f"successfully"}

    except ShapeError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"failed: {e}")
        raise ShapeError(str(e))

def set_verticaltext(
        filepath:str,
        slide_num:int = 0,
        shape_num:int = 0,
        verticaltext_type:str = "Vertical270"
) -> dict[str,Any]:
    try:
        #Create a PPT document
        ppt = Presentation()
        ppt.LoadFromFile(filepath)

        slide = ppt.Slides[slide_num]

        shape = slide.Shapes[shape_num]

        type1 = None
        for type in VerticalTextType:
            if type.name == verticaltext_type:
                type1 = type

        if type1 == None:
            type1 = VerticalTextType.Vertical270

        shape.TextFrame.VerticalTextType = type1
        ppt.SaveToFile(filepath,FileFormat.Pptx2019)
        return {"message": f"successfully"}

    except ShapeError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"failed: {e}")
        raise ShapeError(str(e))
    
def set_text_color(
        filepath:str,
        slide_num:int = 0,
        shape_num:int = 0,
        color: str = None
) -> dict[str,Any]:
    try:
        #Create a PPT document
        ppt = Presentation()
        ppt.LoadFromFile(filepath)

        slide = ppt.Slides[slide_num]

        shape = slide.Shapes[shape_num]

        if color != None:
            shape.TextFrame.TextRange.Fill.FillType = FillFormatType.Solid
            if color.startswith('#'):
                color = color[1:]
            if len(color) == 6:
                r = int(color[0:2], 16)
                g = int(color[2:4], 16)
                b = int(color[4:6], 16)
                shape.TextFrame.TextRange.Fill.SolidColor.Color = Color.FromRgb(r, g, b)
        
        ppt.SaveToFile(filepath,FileFormat.Pptx2019)
        return {"message": f"successfully"}

    except ShapeError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"failed: {e}")
        raise ShapeError(str(e))