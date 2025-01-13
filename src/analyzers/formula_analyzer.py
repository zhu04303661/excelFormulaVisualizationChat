"""Excel公式分析器，用于分析公式依赖关系和构建计算树"""

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from ..utils.cell_utils import is_yellow_cell, get_cell_address
import re

class FormulaAnalyzer:
    def __init__(self, workbook):
        self.workbook = workbook

    def build_formula_tree(self, formulas):
        """
        为保持向后兼容性提供的别名方法，调用analyze_formulas
        
        Args:
            formulas: 公式信息列表
            
        Returns:
            list: 包含简化依赖信息的公式列表
        """
        return self.analyze_formulas(formulas)

    def analyze_formulas(self, formulas):
        """
        分析公式的直接依赖关系
        
        Args:
            formulas: 公式信息列表
            
        Returns:
            list: 包含简化依赖信息的公式列表
        """
        simplified_formulas = []
        
        for formula in formulas:
            # 获取基本信息
            sheet = formula['工作表']
            cell = formula['单元格']
            formula_text = formula['计算方法']
            
            # 获取直接依赖
            deps = [dep for dep in formula['基础单元格依赖'].split(', ') if dep]
            
            # 获取依赖单元格的值
            dep_values = {}
            for dep in deps:
                value = self._get_cell_value(dep)
                dep_values[dep] = value
            
            # 创建简化的公式信息
            simplified_formula = {
                '层级': '基础公式' if len(deps) == 1 else f'第{len(deps)}层',
                '表格名称': sheet,
                '工作表': sheet,
                '单元格': cell,
                '行标题': formula.get('行标题', ''),
                '列标题': formula.get('列标题', ''),
                '计算方法': formula_text,
                '基础单元格依赖': formula['基础单元格依赖'],
                '简化计算表达式': self._simplify_formula(formula_text),
                '变量表达式': formula.get('变量表达式', ''),
                '计算结果': self._get_cell_value(f"{sheet}!{cell}"),
                '直接依赖': {
                    dep: {
                        '值': dep_values[dep],
                        '是否为基础数据': self._is_yellow_cell_ref(dep)
                    } for dep in deps
                }
            }
            
            simplified_formulas.append(simplified_formula)
        
        return simplified_formulas

    def save_analysis(self, simplified_formulas, output_file):
        """
        保存分析结果到Excel文件
        
        Args:
            simplified_formulas: 简化后的公式列表
            output_file: 输出文件路径
        """
        import pandas as pd
        
        # 创建数据框
        data = []
        for formula in simplified_formulas:
            data.append({
                '层级': formula.get('层级', ''),
                '表格名称': formula.get('表格名称', ''),
                '工作表': formula.get('工作表', ''),
                '单元格': formula.get('单元格', ''),
                '行标题': formula.get('行标题', ''),
                '列标题': formula.get('列标题', ''),
                '计算方法': formula.get('计算方法', ''),
                '基础单元格依赖': formula.get('基础单元格依赖', ''),
                '简化计算表达式': formula.get('简化计算表达式', ''),
                '变量表达式': formula.get('变量表达式', '')
            })
        
        # 指定列的顺序
        columns = [
            '层级', '表格名称', '工作表', '单元格', '行标题', '列标题',
            '计算方法', '基础单元格依赖', '简化计算表达式', '变量表达式'
        ]
        
        # 创建DataFrame并指定列顺序
        df = pd.DataFrame(data, columns=columns)
        
        # 保存到Excel
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='公式分析')
            
            # 获取工作表
            ws = writer.sheets['公式分析']
            
            # 设置列宽
            column_widths = {
                'A': 12,  # 层级
                'B': 15,  # 表格名称
                'C': 15,  # 工作表
                'D': 12,  # 单元格
                'E': 20,  # 行标题
                'F': 20,  # 列标题
                'G': 40,  # 公式
                'H': 40,  # 基础单元格依赖
                'I': 30,  # 简化计算表达式
                'J': 40,  # 变量表达式
            }
            
            # 应用列宽
            for col, width in column_widths.items():
                ws.column_dimensions[col].width = width

    def _get_cell_value(self, cell_ref):
        """获取单元格的值"""
        try:
            sheet, addr = cell_ref.split('!')
            cell = self.workbook[sheet][addr]
            return str(cell.value) if cell.value is not None else ''
        except:
            return ''

    def _is_yellow_cell_ref(self, cell_ref):
        """检查单元格是否为基础数据（黄色背景）"""
        try:
            sheet, addr = cell_ref.split('!')
            cell = self.workbook[sheet][addr]
            if cell.fill.start_color.type == 'theme':
                return False
            
            rgb = cell.fill.start_color.rgb
            if not rgb:
                return False
            
            r = int(rgb[2:4], 16)
            g = int(rgb[4:6], 16)
            b = int(rgb[6:8], 16)
            
            return r > 200 and g > 200 and b < 100
        except:
            return False 

    def _simplify_formula(self, formula):
        """简化公式表达式"""
        if '=' in formula:
            formula = formula.split('=')[1].strip()
        if '(' in formula:
            # 提取函数名和主要参数
            parts = formula.split('(')
            func_name = parts[0]
            return f"{func_name}(...)"
        return formula[:30] + '...' if len(formula) > 30 else formula

    def _build_trace_path(self, deps, dep_values):
        """构建追溯路径"""
        path_parts = []
        for dep in deps:
            dep_type = '基础数据' if self._is_yellow_cell_ref(dep) else '计算方法'
            value = dep_values[dep]
            path_parts.append(f"{dep}[{dep_type}]={value}")
        return ' -> '.join(path_parts) 

    def generate_formula_tree_html(self, formulas, output_file):
        """
        生成公式依赖关系的HTML可视化文件
        
        Args:
            formulas: 公式信息列表
            output_file: 输出文件路径
            
        Returns:
            str: 生成的HTML文件路径
        """
        # 生成HTML文件路径
        html_file = output_file.replace('.xlsx', '.html')
        
        # 生成HTML内容
        html_content = [
            '<!DOCTYPE html>',
            '<html>',
            '<head>',
            '<meta charset="UTF-8">',
            '<title>公式依赖关系图</title>',
            '<script src="https://cdn.jsdelivr.net/npm/mermaid@9.1.3/dist/mermaid.min.js"></script>',
            '<style>',
            'body { font-family: Arial, sans-serif; margin: 20px; }',
            '.container { display: flex; }',
            '.mermaid { flex: 1; }',
            '.info-panel { width: 300px; padding: 20px; }',
            '</style>',
            '</head>',
            '<body>',
            '<div class="container">',
            '<div class="mermaid">'
        ]
        
        # 生成Mermaid图表内容
        mermaid_content = ['graph TB']
        nodes = {}
        node_counter = 0
        
        # 添加节点
        for formula in formulas:
            node_id = f'n{node_counter}'
            node_counter += 1
            cell_ref = f"{formula['工作表']}!{formula['单元格']}"
            
            # 节点标签
            label = f"{formula['单元格']}<br/>{self._simplify_formula(formula['计算方法'])}"
            nodes[cell_ref] = node_id
            
            # 添加节点定义
            mermaid_content.append(f'    {node_id}["{label}"]')
            
            # 添加节点样式
            if formula['层级'] == '基础公式':
                mermaid_content.append(f'    class {node_id} basic')
            else:
                mermaid_content.append(f'    class {node_id} formula')
        
        # 添加依赖关系连接
        for formula in formulas:
            if formula['基础单元格依赖']:
                target_node = nodes[f"{formula['工作表']}!{formula['单元格']}"]
                for dep in formula['基础单元格依赖'].split(', '):
                    if dep in nodes:  # 只连接到已存在的节点
                        source_node = nodes[dep]
                        mermaid_content.append(f'    {source_node} --> {target_node}')
        
        # 添加样式定义
        mermaid_content.insert(1, '    classDef basic fill:#e1f5fe,stroke:#333')
        mermaid_content.insert(1, '    classDef formula fill:#f9f9f9,stroke:#333')
        
        # 将Mermaid内容添加到HTML
        html_content.extend(mermaid_content)
        
        # 添加HTML尾部
        html_content.extend([
            '</div>',
            '<div class="info-panel">',
            '<h3>图例说明</h3>',
            '<div style="margin: 10px 0;">',
            '<div style="background: #f9f9f9; border: 1px solid #333; padding: 5px; margin: 5px 0;">复合公式</div>',
            '<div style="background: #e1f5fe; border: 1px solid #333; padding: 5px; margin: 5px 0;">基础公式</div>',
            '</div>',
            '</div>',
            '</div>',
            '<script>',
            'mermaid.initialize({',
            '    startOnLoad: true,',
            '    theme: "default",',
            '    flowchart: {',
            '        useMaxWidth: true,',
            '        htmlLabels: true,',
            '        curve: "basis",',
            '        nodeSpacing: 50,',
            '        rankSpacing: 50',
            '    }',
            '});',
            '</script>',
            '</body>',
            '</html>'
        ])
        
        # 保存HTML文件
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(html_content))
        
        return html_file 

    def synthesize_formula(self, output_cell_ref, formulas, input_cells):
        """
        合成公式，将输出单元格的公式追溯到只使用输入单元格
        
        Args:
            output_cell_ref: 输出单元格引用（如'Sheet1!A1'）
            formulas: 所有公式的列表
            input_cells: 输入单元格列表
            
        Returns:
            str: 合成后的公式
        """
        # 创建输入单元格集合，用于快速查找
        input_cell_refs = {f"{cell['工作表']}!{cell['单元格']}" for cell in input_cells}
        
        # 创建公式字典，用于快速查找
        formula_dict = {f"{f['工作表']}!{f['单元格']}": f for f in formulas}
        
        def substitute_formula(cell_ref, visited=None):
            if visited is None:
                visited = set()
                
            # 防止循环引用
            if cell_ref in visited:
                return None
            visited.add(cell_ref)
            
            # 如果是输入单元格，返回其变量名
            if cell_ref in input_cell_refs:
                input_cell = next(cell for cell in input_cells 
                                if f"{cell['工作表']}!{cell['单元格']}" == cell_ref)
                return input_cell['标题组合']
            
            # 如果不是公式单元格，返回None
            if cell_ref not in formula_dict:
                return None
                
            formula = formula_dict[cell_ref]
            current_formula = formula['计算方法']
            
            # 查找公式中的所有单元格引用
            cell_refs = re.findall(r"([^+\-*/(),\s!]+![A-Z]+[0-9]+)", current_formula)
            
            # 替换每个引用
            for ref in cell_refs:
                substituted = substitute_formula(ref, visited)
                if substituted:
                    current_formula = current_formula.replace(ref, f"({substituted})")
                    
            return current_formula
            
        # 开始合成公式
        result = substitute_formula(output_cell_ref)
        return result if result else "无法合成公式" 