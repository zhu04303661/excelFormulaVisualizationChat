"""Excel公式分析器，用于分析公式依赖关系和构建计算树"""

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from ..utils.cell_utils import is_yellow_cell, get_cell_address
import re

class FormulaAnalyzer:
    def __init__(self, workbook):
        self.workbook = workbook