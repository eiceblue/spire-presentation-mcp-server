import logging
from typing import Any

from spire.presentation import *

from .exceptions import SmartArtError

logger = logging.getLogger(__name__)

def create_smartart(
        filepath:str,
        slide_num:int,
        x:float = 0,
        y:float = 0,
        width:float = 200,
        height:float = 200,
        layout_type:str = "Gear"
) -> dict[str,Any]:
    try:
        #Create a PPT document
        ppt = Presentation()
        ppt.LoadFromFile(filepath)

        slide = ppt.Slides[slide_num]

        type1 = None
        for type in SmartArtLayoutType:
            if type.name == layout_type:
                type1 = type

        if type1 == None:
            type1 = SmartArtLayoutType.Gear

        slide.Shapes.AppendSmartArt(x,y,width,height,type1)

        #Save the document
        ppt.SaveToFile(filepath,FileFormat.Pptx2019)
        return {"message": f"add successfully"}

    except SmartArtError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"add failed: {e}")
        raise SmartArtError(str(e))