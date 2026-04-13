from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter

def make_directory_catalog(folder: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = '目录'
    ws.merge_cells('A1:F1')
    ws['A1'] = '目录清单'
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

    headers = ['序号', '名称', '文件格式（zip、rar、word、pdf、ppt等）', '文件大小', '是否经分管领导审核', '备注']
    ws.append(headers)

    thin = Side(border_style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    alignment = Alignment(horizontal='center', vertical='center')
    font_text = Font(name='黑体', size=14)
    font_title = Font(name="黑体", size=20, bold=True)

    # data rows
    index = 1
    for file in sorted(folder.iterdir()): # sorted按照文件名排序，iterdir()返回目录下的所有文件和文件夹的Path对象, iterdir()默认只返回当前目录下的内容，不会递归进入子目录。
        if file.is_file():
            name = file.name
            suffix = file.suffix.lstrip('.').lower() or '' # lstrip全称是去除左边的字符,lower()转换为小写，如果没有后缀则返回空字符串
            size = file.stat().st_size # stat()函数返回一个包含文件状态信息的对象，st_size属性表示文件的大小（以字节为单位）
            # 根据文件大小自动显示单位
            if size < 1024:
                size_str = f"{size}B"
            elif size < 1024 * 1024:
                size_str = f"{round(size / 1024)}kb"
            else:
                size_str = f"{round(size / (1024 * 1024), 2)}mb" # 2表示保留两位小数
            ws.append([index, name, suffix, size_str, '是', ''])
            index += 1
        if file.is_dir():
            name = file.name # name返回文件夹的名称
            ws.append([index, name, '文件夹', '', '是', ''])
            index += 1

    # style header cells
    max_row = ws.max_row
    max_col = ws.max_column
    title_cell = ws.cell(1,1)
    title_cell.font = font_title
    title_cell.border = border
    for c in range(1, max_col + 1):
        header_cell = ws.cell(2, c)
        header_cell.font = font_title
        header_cell.border = border
        header_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    # all content cells
    for r in range(3, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = font_text
            cell.border = border
            cell.alignment = alignment
    # 所有列设置固定宽度22 
    for c in range(1, max_col + 1):
        col_letter = get_column_letter(c)
        ws.column_dimensions[col_letter].width = 22  # 固定列宽为22, column_dimensions是openpyxl中用于设置列属性的对象，get_column_letter函数将列索引转换为Excel列字母，例如1对应A，2对应B，以此类推。
        # 至少设置第三行的格式
        cell = ws.cell(row=3, column=c)
        cell.font = font_text
        cell.border = border
        cell.alignment = alignment

    # # auto column width
    # for c in range(1, max_col + 1):
    #     col_letter = ws.cell(row=1, column=c).column_letter
    #     max_length = 0
    #     for r in range(1, max_row + 1):
    #         value = ws.cell(row=r, column=c).value
    #         if value is None:
    #             continue
    #         max_length = max(max_length, len(str(value)))
    #     ws.column_dimensions[col_letter].width = max_length + 2

    target = folder / '0.目录.xlsx'
    wb.save(str(target))
    return target

def main():
    root_H = Path('H:/')
    first_level_folder = root_H / '西安-近5年中涉及调用原子能力的合同的资料'
    if not first_level_folder.exists():
        print(f'未找到一级目录: {first_level_folder}')
        return
    # 生成一级目录文件
    make_directory_catalog(first_level_folder)
    for second_level_folder in first_level_folder.iterdir(): 
        # 跳过新生成的一级目录文件
        if second_level_folder.is_dir():
            # 生成二级目录文件
            print(f'正在处理目录: {second_level_folder}')
            make_directory_catalog(second_level_folder)
            for third_level_folder in second_level_folder.iterdir():
                # 跳过新生成的二级目录文件
                if third_level_folder.is_dir():
                    # 生成三级目录文件
                    print(f'正在处理目录: {third_level_folder}')
                    make_directory_catalog(third_level_folder)
    print('各级目录生成完毕！')

if __name__ == '__main__':
    main()