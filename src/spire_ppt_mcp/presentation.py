import logging
from pathlib import Path
from typing import Any

from spire.presentation import *

from .exceptions import PresentationError

logger = logging.getLogger(__name__)

def create_presentation(filepath:str) -> dict[str]:
    """Create a new presentation with optional custom ppt name"""
    try:
        ppt = Presentation()

        save_path = Path(filepath)
        save_path.parent.mkdir(parents=True,exist_ok=True)

        ppt.SaveToFile(str(save_path),FileFormat.Pptx2019)
        return{
            "message":f"Created Presentation:{filepath}",
            "presentation":ppt
        }
    except Exception as e:
        logger.error(f"Failed to create presentation: {e}")
        raise PresentationError(f"Failed to create presentation: {e!s}")
    
def get_or_create_presentation(filepath: str) -> Presentation:
    """Get existing presentation or create new one if it doesn't exist"""
    try:
        ppt = Presentation()
        if Path(filepath).exists():
            # 加载已有的 PPT 文件
            ppt.LoadFromFile(filepath)
        else:
            # 创建新的 PPT，并确保目录存在
            ppt = create_presentation(filepath)
        return ppt
    except Exception as e:
        logger.error(f"Failed to get or create presentation: {e}")
        raise PresentationError(f"Failed to get or create presentation: {e!s}")
    
# def create_slide(filepath: str) -> dict:
#     """
#     Create a new slide in the presentation with the given title if it doesn't exist.
#     """
#     try:
#         ppt = Presentation()
#         ppt.LoadFromFile(filepath)

#         # 创建新的幻灯片
#         ppt.Slides.Append()
#         ppt.SaveToFile(filepath)
#         return {"message": f"Slide created successfully"}

#     except PresentationError as e:
#         logger.error(str(e))
#         raise
#     except Exception as e:
#         logger.error(f"Failed to create slide: {e}")
#         raise PresentationError(str(e))