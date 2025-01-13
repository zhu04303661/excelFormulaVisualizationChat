"""Excel公式分析主程序"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from src.extractors.formula_extractor import FormulaExtractor
import traceback
#import ipdb
import os
import argparse
import traceback
from openpyxl import load_workbook
from src.extractors.formula_extractor import FormulaExtractor


def save_input_cells_to_text(input_cells, output_file='input_cells.txt'):
    """
    将输入单元格信息保存到文本文件
    
    Args:
        input_cells (list): 输入单元格信息列表
        output_file (str): 输出文件路径，默认为'input_cells.txt'
    """
    if not input_cells:
        print("没有找到输入单元格")
        return
        
    output_text = []
    current_table = None
    
    output_text.append("输入单元格列表：\n")
    #ipdb.set_trace()
    for cell_info in input_cells:
        # 当遇到新的表格名称时，添加分隔行
        if cell_info['表格名称'] != current_table:
            current_table = cell_info['表格名称']
            separator = f"\n{'='*80}\n表格名称: {current_table}\n{'='*80}"
            output_text.append(separator)
        
        # 格式化输出每个输入单元格的信息
        cell_info_text = f"""
输入变量名称: {cell_info['标题组合']}
说明：已输入值: {cell_info['当前值']}  位置: {cell_info['工作表']}!{cell_info['单元格']}
"""
        #     cell_info_text = f"""{cell_info['标题组合']}"""
        output_text.append(cell_info_text)
    
    # 保存到文本文件
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(output_text))
    
    print(f"\n找到 {len(input_cells)} 个输入单元格")
    print(f"输入单元格信息已保存到: {output_file}")

def save_output_cells_to_text(output_cells, output_file='output_cells.txt'):
    """
    将输出单元格信息保存到文本文件
    
    Args:
        output_cells (list): 输出单元格信息列表
        output_file (str): 输出文件路径，默认为'output_cells.txt'
    """
    if not output_cells:
        print("没有找到输出单元格")
        return
        
    output_text = []
    current_table = None
    
    output_text.append("输出单元格列表：\n")
    
    for cell_info in output_cells:
        # 当遇到新的表格名称时，添加分隔行
        if cell_info['表格名称'] != current_table:
            current_table = cell_info['表格名称']
            separator = f"\n{'='*80}\n表格名称: {current_table}\n{'='*80}"
            output_text.append(separator)
        
        # 格式化输出每个输出单元格的信息
        cell_info_text = f"""
输出变量名称: {cell_info['标题组合']}
说明：计算结果: {cell_info['当前值']}  位置: {cell_info['工作表']}!{cell_info['单元格']}
{'-'*80}"""
        output_text.append(cell_info_text)
    
    # 保存到文本文件
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(output_text))
    
    print(f"\n找到 {len(output_cells)} 个输出单元格")
    print(f"输出单元格信息已保存到: {output_file}")

def process_excel_formulas(input_file, output_file):
    """
    处理Excel文件中的公式
    
    Args:
        input_file (str): 输入Excel文件路径
        output_file (str): 输出Excel文件路径
    """
    try:
        # 加载工作簿
        print('正在加载Excel文件...')
        workbook = load_workbook(input_file, data_only=False)
        
        # 提取公式
        print('正在提取公式...')
        formula_extractor = FormulaExtractor(workbook, input_file)
        
        #if formulas:
        if formula_extractor:
            
            # 保存输入单元格信息到文本文件
            save_input_cells_to_text(formula_extractor.input_cells)
            
            # 保存输出单元格信息到文本文件
            save_output_cells_to_text(formula_extractor.output_cells)
            
            
            formula_extractor._analyze_formula_dependencies(formula_extractor.output_cells)
            
        else:
            print('未找到任何公式')
            
    except Exception as e:
        print(f'处理过程出现错误: {str(e)}')
        print('\n详细错误信息:')
        print(traceback.format_exc())
    
            
def main():
    """
    主函数，处理命令行参数并执行公式处理
    """
    # 设置命令行参数ˆ
    parser = argparse.ArgumentParser(description='Excel公式分析工具')
    parser.add_argument('--input_file',  
                        default='input.xlsx',
                        help='输入Excel文件路径')
    parser.add_argument('--output', '-o', 
                       default='formula_analysis_result.xlsx',
                       help='输出Excel文件路径 (默认: formula_analysis_result.xlsx)')
    
    # 解析命令行参数
    args = parser.parse_args()
    

    # 检查输入文件是否存在
    if not os.path.exists(args.input_file):
        print(f'错误: 输入文件不存在: {args.input_file}')
        return
    
    try:
        # 处理Excel公式
        process_excel_formulas(args.input_file, args.output)
    except Exception as e:
        print(f'程序执行出错: {str(e)}')
        print('\n详细错误信息:')
        print(traceback.format_exc())

if __name__ == '__main__':
    main()