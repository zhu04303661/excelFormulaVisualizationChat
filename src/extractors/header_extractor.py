"""Excel标题提取器，用于提取行列标题"""

from ..utils.cell_utils import is_numeric, is_yellow_cell
#import ipdb;

class HeaderExtractor:
    def __init__(self, worksheet):
        """
        初始化标题提取器
        
        Args:
            worksheet: openpyxl工作表对象
        """
        self.ws = worksheet
        self.special_keywords = ['万元', '元', '%', '百分比', '比例']
        
    def find_numeric_header_title(self, row, col):
        """
        当列标题为"无标题"时的补充检索
        查找数字序号，并在其所在行寻找非数字标题进行组合
        
        Args:
            row (int): 当前行号
            col (int): 当前列号
            
        Returns:
            str: 组合后的标题，如果未找到则返回None
        """
        # 向上查找数字序号
        for r in range(row-1, 0, -1):
            cell = self.ws.cell(row=r, column=col)
            # 检查是否是合并单元格
            if any(cell.coordinate in merged_range for merged_range in self.ws.merged_cells.ranges):
                break  # 遇到合并单元格就停止搜索
            
            value = cell.value
            if value and str(value).strip() != '0' and not (isinstance(value, str) and value.startswith('=')):
                current_value = str(value).strip()
                if is_numeric(current_value):
                    # 检查上一个单元格
                    prev_cell = self.ws.cell(row=r-1, column=col) if r > 1 else None
                    is_prev_merged = prev_cell and any(prev_cell.coordinate in merged_range for merged_range in self.ws.merged_cells.ranges)
                    
                    # 如果上一个单元格是合并单元格或者已经到达顶部，才向左查找非数字标题
                    if is_prev_merged or r == 1:
                        # 在该数字所在行向左查找非数字标题
                        for c in range(col-1, 0, -1):
                            left_cell = self.ws.cell(row=r, column=c)
                            # 检查是否是合并单元格
                            # if any(left_cell.coordinate in merged_range for merged_range in self.ws.merged_cells.ranges):
                            #     break  # 遇到合并单元格就停止搜索
                            
                            left_value = left_cell.value
                            if left_value and str(left_value).strip() != '0' and not (isinstance(left_value, str) and left_value.startswith('=')):
                                if not is_numeric(str(left_value)):
                                    return f"{str(left_value).strip()}_{current_value}"
                                # 如果是数字，继续搜索
                                continue
        return None
        
    def find_nearest_header(self, row, col, direction='both'):
        """
        查找最近的标题，遇到合并单元格时停止搜索
        跳过数字继续搜索，直到遇到合并单元格
        
        Args:
            row (int): 当前行号
            col (int): 当前列号
            direction (str): 搜索方向，'both'表示同时搜索行列标题，'row'仅搜索行标题，'column'仅搜索列标题
            
        Returns:
            tuple: (行标题, 列标题)，如果未找到则对应位置为None
        """
        row_header = []
        col_header = []
        
        targert_cell = self.ws.cell(row=row, column=col)
        
        # 向左查找最近的行标题
        found_row_header_in_first_col = False
        if direction in ['both', 'row']:
            for c in range(col-1, 0, -1):
                cell = self.ws.cell(row=row, column=c)               
                value = cell.value
                # 跳过空值、公式和数字
                if value and str(value).strip() != '0' and not (isinstance(value, str) and value.startswith('=')):
                    current_value = str(value).strip()
                    if not is_numeric(current_value):
                        row_header.insert(0, current_value)
                        if c == col-1:  # 如果在第一列就找到了非数字标题
                            found_row_header_in_first_col = True
                        # 如果当前值不是特殊关键词，就停止继续查找
                        if current_value not in self.special_keywords:
                            break
                    # 如果是数字，继续搜索
                    continue

        # 检查右侧列是否为空
        right_col_empty = True
        for r in range(1, self.ws.max_row + 1):
            right_cell = self.ws.cell(row=r, column=col+1)
            if right_cell.value:
                right_col_empty = False
                break
  
        # 只有当右侧列不为空，或者没有在第一列找到行标题时，才向上查找列标题
        if direction in ['both', 'column'] and (not right_col_empty or not found_row_header_in_first_col):
            for r in range(row-1, 0, -1):
                cell = self.ws.cell(row=r, column=col)
                # 检查是否是合并单元格
                if any(cell.coordinate in merged_range for merged_range in self.ws.merged_cells.ranges):
                    break  # 遇到合并单元格就停止搜索
                
                value = cell.value
                # 跳过空值、公式和数字
                if value and str(value).strip() != '0' and not (isinstance(value, str) and value.startswith('=')):
                    current_value = str(value).strip()
                    if not is_numeric(current_value):
                        col_header.insert(0, current_value)
                        # 如果当前值不是特殊关键词，就停止继续查找
                        if current_value not in self.special_keywords:
                            break
                    # 如果是数字，继续搜索
                    continue

        # 组合标题
        row_header_text = '_'.join(row_header) if row_header else None
        col_header_text = '_'.join(col_header) if col_header else None
        
        # 只有当右侧列不为空，或者没有在第一列找到行标题时，才尝试补充检索列标题
        if not col_header_text and (not right_col_empty or not found_row_header_in_first_col):
            col_header_text = self.find_numeric_header_title(row, col)
        
        return row_header_text, col_header_text 