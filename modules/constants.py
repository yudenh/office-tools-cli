from enum import Enum


class HandleResult(Enum):
    """
    处理结果"
    """

    Silent = 0
    Created = 1
    Overrided = 2
    Skiped = 3
    Processed = 4
