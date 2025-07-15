import logging
from typing import Any

from spire.presentation import *

from .exceptions import ChartError

logger = logging.getLogger(__name__)

def add_chart(
        filepath:str,
        slide_num:int,
        x:float = 0,
        y:float = 0,
        width:float = 200,
        height:float = 200,
        chart_type:str = "Pie"
) -> dict[str,Any]:
    try:
        #Create a PPT document
        ppt = Presentation()
        ppt.LoadFromFile(filepath)

        slide = ppt.Slides[slide_num]

        rect = RectangleF.FromLTRB (x, y, width + x, height + y)

        type1 = None
        for type in ChartType:
            if type.name == chart_type:
                type1 = type

        if type1 == None:
            type1 = ChartType.Pie

        slide.Shapes.AppendChart(type1,rect)

        #Save the document
        ppt.SaveToFile(filepath,FileFormat.Pptx2019)
        return {"message": f"add successfully"}

    except ChartError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"add failed: {e}")
        raise ChartError(str(e))