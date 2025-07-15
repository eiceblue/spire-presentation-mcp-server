import logging
from typing import Any

from spire.presentation import *

from .exceptions import TableError

logger = logging.getLogger(__name__)

def create_table(
        filepath:str,
        slide_num:int,
        x:float = 0,
        y:float = 0,
        widths:List[float] = [50,50],
        heights:List[float] = [20,20]
) -> dict[str,Any]:
    try:
        #Create a PPT document
        ppt = Presentation()
        ppt.LoadFromFile(filepath)

        slide = ppt.Slides[slide_num]
        print(widths)
        print(heights)
        slide.Shapes.AppendTable(x,y,widths,heights)

        #Save the document
        ppt.SaveToFile(filepath,FileFormat.Pptx2019)
        return {"message": f"add successfully"}

    except TableError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"add failed: {e}")
        raise TableError(str(e))

def add_text_table(
        filepath:str,
        slide_num:int,
        shape_num:int,
        data_str:List[str] = ["", "", "", ""]
) -> dict[str,Any]:
    try:
        #Create a PPT document
        ppt = Presentation()
        ppt.LoadFromFile(filepath)

        slide = ppt.Slides[slide_num]

        table = slide.Shapes[shape_num]

        if not isinstance(table, ITable):
            return {"success": False, "message": "The specified shape is not a table."}

        row_count = table.TableRows.Count
        col_count = table.TableRows[0].Count if row_count > 0 else 0

        data_2d = []

        for i in range(0, len(data_str), col_count):
            row = data_str[i:i + col_count]
            # 如果最后一行不够，补空字符串 ""
            if len(row) < col_count:
                row += [""] * (col_count - len(row))
            data_2d.append(row)

        if isinstance(table,ITable):
            for i in range(0,table.TableRows.Count):
                for j in range(0,table.TableRows[i].Count):
                    table[j,i].TextFrame.Text = data_2d[i][j]

        #Save the document
        ppt.SaveToFile(filepath,FileFormat.Pptx2019)
        return {
            "success": True,
            "message": "Table content initialized successfully."
        }

    except TableError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"add failed: {e}")
        raise TableError(str(e))

