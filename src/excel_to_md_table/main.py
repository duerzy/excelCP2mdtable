import pyperclip
import argparse
import re

def excel_to_md_table(content):
    # 将Excel内容转换为Markdown表格的逻辑
    
    # 首先处理引号包裹的内容，将其中的换行替换为特殊标记
    content = re.sub(r'"([^"]*)"', lambda m: m.group(0).replace('\n', 'NEWLINE'), content)
    
    # 分割行
    rows = content.strip().split('\n')
    rows = [row for row in rows if row.strip()]

    if not rows:
        return "无有效内容"

    # 处理表头
    headers = process_row(rows[0])
    md_table = "| " + " | ".join(headers) + " |\n"
    md_table += "|" + "---|" * len(headers) + "\n"

    # 处理数据行
    for row in rows[1:]:
        cells = process_row(row)
        md_table += "| " + " | ".join(cells) + " |\n"

    return md_table

def process_row(row):
    # 使用制表符分割单元格
    cells = row.split('\t')
    
    # 处理每个单元格
    processed_cells = []
    for cell in cells:
        cell = cell.strip()
        if cell.startswith('"') and cell.endswith('"'):
            # 处理引号内的内容
            processed_cell = cell[1:-1].replace('NEWLINE', '<br>')
        else:
            processed_cell = cell
        processed_cells.append(processed_cell)
    
    return processed_cells

def main():
    parser = argparse.ArgumentParser(description="将Excel内容转换为Markdown表格")
    parser.add_argument("-i", "--input", help="输入文件路径")
    parser.add_argument("-o", "--output", help="输出文件路径")
    args = parser.parse_args()

    if args.input:
        with open(args.input, 'r', encoding='utf-8') as f:
            content = f.read()
    else:
        content = pyperclip.paste()

    md_table = excel_to_md_table(content)

    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(md_table)
    else:
        pyperclip.copy(md_table)
        print("Markdown表格已复制到剪贴板")

if __name__ == "__main__":
    main()