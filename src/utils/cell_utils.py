"""Excel单元格工具类，提供单元格操作的通用功能"""

from openpyxl.utils import get_column_letter, column_index_from_string

def is_numeric(value):
    """
    检查值是否为数字（包括整数、小数、科学计数法）
    
    Args:
        value: 要检查的值
        
    Returns:
        bool: 是否为数字
    """
    if value is None:
        return False
    try:
        float(str(value))
        return True
    except ValueError:
        return False

def get_cell_address(sheet_name, cell_ref):
    """
    获取完整的单元格地址
    
    Args:
        sheet_name (str): 工作表名称
        cell_ref (str): 单元格引用（如A1）
        
    Returns:
        str: 完整的单元格地址（如Sheet1!A1）
    """
    return f"{sheet_name}!{cell_ref}"

def parse_cell_reference(cell_ref):
    """
    解析单元格引用，返回工作表名和单元格地址
    
    Args:
        cell_ref (str): 单元格引用（如Sheet1!A1或A1）
        
    Returns:
        tuple: (工作表名, 单元格地址)，如果没有工作表名则为(None, cell_ref)
    """
    if '!' in cell_ref:
        sheet_name, address = cell_ref.split('!')
        return sheet_name.strip("'"), address
    return None, cell_ref

def get_column_row_from_cell_ref(cell_ref):
    """
    从单元格引用中获取列和行
    
    Args:
        cell_ref (str): 单元格引用（如A1）
        
    Returns:
        tuple: (列字母, 行号)
    """
    col_str = ''.join(filter(str.isalpha, cell_ref))
    row_num = int(''.join(filter(str.isdigit, cell_ref)))
    return col_str, row_num

def is_valid_cell_reference(cell_ref):
    """
    检查单元格引用是否有效
    
    Args:
        cell_ref (str): 单元格引用（如A1）
        
    Returns:
        bool: 是否为有效的单元格引用
    """
    try:
        col_str, row_num = get_column_row_from_cell_ref(cell_ref)
        column_index_from_string(col_str)  # 验证列字母是否有效
        return True
    except:
        return False

def is_yellow_cell(cell):
    """
    检查单元格是否有黄色背景
    
    Args:
        cell: 单元格对象
        
    Returns:
        bool: 是否是黄色背景
    """
    if cell.fill.start_color.index:
        # 检查是否是黄色背景（可能需要根据实际使用的黄色色值调整）
        yellow_colors = ['FFFF00', 'FFFFE0', 'FFFFD7', 'FFFFF0', 'FFFFFF00','FFFFFFF0']  # 可能的黄色色值
        return cell.fill.start_color.index in yellow_colors
    return False



def is_blue_cell(cell):
    """
    检查单元格是否为蓝色背景
    
    Args:
        cell: openpyxl单元格对象
        
    Returns:
        bool: 是否为蓝色背景
    """
    return cell.fill.start_color.rgb == "FF00B0F0" 

