#!/usr/bin/env python3
"""
报表模板公式检查工具

功能：
1. 统计各工作表的公式数量
2. 检查是否还有通配符公式
3. 验证精确匹配公式数量
4. 生成检查报告

使用方法：
    python3 check_formulas.py [文件名]
"""

import sys
from openpyxl import load_workbook

def check_formulas(filepath):
    """检查 Excel 文件中的公式"""
    print("=" * 70)
    print(f"📊 公式检查报告")
    print(f"文件: {filepath}")
    print("=" * 70)

    wb = load_workbook(filepath, data_only=False)

    results = []
    total_wildcard = 0
    total_trim = 0
    total_formulas = 0

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        wildcard = 0
        trim = 0
        formulas = 0
        wildcard_cells = []

        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                    formulas += 1
                    formula = cell.value

                    # 检查通配符 (排除 SEARCH 函数中的合法通配符)
                    if ('MATCH("*' in formula or '&"*"' in formula):
                        wildcard += 1
                        wildcard_cells.append(cell.coordinate)

                    # 检查精确匹配
                    if 'TRIM(' in formula and '!$B$' in formula:
                        trim += 1

        if formulas > 0:
            status = "✅" if wildcard == 0 else "⚠️"
            results.append({
                'name': sheet_name,
                'formulas': formulas,
                'wildcard': wildcard,
                'trim': trim,
                'status': status,
                'wildcard_cells': wildcard_cells[:5]  # 只显示前5个
            })

            total_wildcard += wildcard
            total_trim += trim
            total_formulas += formulas

    wb.close()

    # 打印结果
    print(f"\n{'工作表':<25} {'公式数':>8} {'通配符':>8} {'精确匹配':>8} {'状态':>6}")
    print("-" * 70)

    for r in results:
        print(f"{r['name']:<25} {r['formulas']:>8} {r['wildcard']:>8} {r['trim']:>8} {r['status']:>6}")

        # 显示有问题的单元格
        if r['wildcard_cells']:
            print(f"    └─ 通配符单元格: {', '.join(r['wildcard_cells'])}")

    print("-" * 70)
    print(f"{'总计':<25} {total_formulas:>8} {total_wildcard:>8} {total_trim:>8}")

    # 总结
    print("\n" + "=" * 70)
    if total_wildcard == 0:
        print("✅ 所有查找公式已使用精确匹配!")
    else:
        print(f"⚠️ 发现 {total_wildcard} 个通配符公式需要修改")
    print("=" * 70)

    return total_wildcard == 0


if __name__ == "__main__":
    if len(sys.argv) > 1:
        filepath = sys.argv[1]
    else:
        filepath = "报表模板20260226_xxd.xlsm"

    try:
        check_formulas(filepath)
    except FileNotFoundError:
        print(f"❌ 文件不存在: {filepath}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ 错误: {e}")
        sys.exit(1)
