import openpyxl
import zipfile
import tempfile
import os
import shutil
import warnings
import re
import traceback
import sys

# 屏蔽警告
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


def strip_and_load(xlsx_path):
    print(f"正在进行全域基因修复: {xlsx_path} ...")

    temp_dir = tempfile.mkdtemp()
    temp_path = os.path.join(temp_dir, "full_transplant_" + os.path.basename(xlsx_path))

    try:
        with zipfile.ZipFile(xlsx_path, 'r') as zin:
            with zipfile.ZipFile(temp_path, 'w') as zout:
                for item in zin.infolist():
                    # 1. 物理移除干扰文件 (保持不变)
                    if 'externalLinks' in item.filename: continue
                    if 'calcChain' in item.filename: continue
                    if 'printerSettings' in item.filename: continue
                    if 'vbaProject' in item.filename: continue

                    content = zin.read(item.filename)
                    filename = item.filename

                    # 2. 核心逻辑：对所有 XML 文件进行命名空间“洗白”
                    # 重点关注: workbook.xml, worksheet XMLs, sharedStrings.xml, styles.xml
                    if filename.endswith('.xml') and (
                            filename.endswith('workbook.xml') or
                            'worksheets/sheet' in filename or
                            'sharedStrings.xml' in filename or
                            'styles.xml' in filename
                    ):
                        xml_str = content.decode('utf-8')

                        # === [全域基因编辑] ===
                        # 替换 Strict OOXML (WPS) -> Transitional OOXML (Microsoft)

                        # 1. SpreadsheetML Main
                        xml_str = xml_str.replace(
                            'http://purl.oclc.org/ooxml/spreadsheetml/main',
                            'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                        )
                        # 2. Relationships
                        xml_str = xml_str.replace(
                            'http://purl.oclc.org/ooxml/officeDocument/relationships',
                            'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                        )
                        # 3. Drawing / VML (防止图片引用导致命名空间报错)
                        xml_str = xml_str.replace(
                            'http://purl.oclc.org/ooxml/drawingml/main',
                            'http://schemas.openxmlformats.org/drawingml/2006/main'
                        )

                        # === [特定文件清洗] ===

                        # 针对 workbook.xml: 移除 definedNames 和 externalReferences
                        if filename.endswith('workbook.xml'):
                            xml_str = re.sub(r'<[\w:]*?definedNames.*?>.*?</[\w:]*?definedNames>', '', xml_str,
                                             flags=re.DOTALL)
                            xml_str = re.sub(r'<[\w:]*?definedNames.*?/>', '', xml_str)
                            xml_str = re.sub(r'<[\w:]*?externalReferences.*?>.*?</[\w:]*?externalReferences>', '',
                                             xml_str, flags=re.DOTALL)
                            xml_str = re.sub(r'<[\w:]*?externalReferences.*?/>', '', xml_str)

                        content = xml_str.encode('utf-8')

                    # 3. 针对 .rels 的清洗
                    elif filename.endswith('.rels'):
                        xml_str = content.decode('utf-8')
                        # 替换关系文件的命名空间
                        xml_str = xml_str.replace(
                            'http://purl.oclc.org/ooxml/package/relationships',
                            'http://schemas.openxmlformats.org/package/2006/relationships'
                        )
                        # 移除 externalLink
                        if 'workbook.xml.rels' in filename:
                            xml_str = re.sub(r'<Relationship[^>]*?externalLink[^>]*?/>', '', xml_str)
                        content = xml_str.encode('utf-8')

                    zout.writestr(item, content)

        print("全域修复完毕，加载中...")

        # 尝试标准加载
        # 这里建议先用 data_only=True (看数值)，如果全是 None，再换 data_only=False (看公式)
        try:
            wb = openpyxl.load_workbook(temp_path, data_only=True, read_only=False, keep_links=False)
            return wb
        except Exception as e:
            print(f"标准模式加载异常 ({e})，切换只读模式...")
            wb = openpyxl.load_workbook(temp_path, data_only=True, read_only=True, keep_links=False)
            return wb

    except Exception as e:
        print(f"!!! 处理失败: {e}")
        traceback.print_exc()
        return None


def map_excel_structure(filepath):
    wb = strip_and_load(filepath)
    if not wb: return

    print(f"\n=== EXCEL STRUCTURE MAP (Full Data Check) ===\n")

    try:
        sheet_names = wb.sheetnames
        print(f"[Sheet List]: {sheet_names}\n")

        # 挑选几个重点 Sheet 进行采样
        # 优先看 '主要指标' 和 '关键假设(必填)'
        target_sheets = sheet_names[:5]

        for sheet_name in target_sheets:
            ws = wb[sheet_name]
            print(f"--- SHEET: {sheet_name} ---")

            # 强制重算维度（如果 xml 没写维度）
            if hasattr(ws, 'calculate_dimension'):
                try:
                    ws.calculate_dimension()
                except:
                    pass

            try:
                max_col = ws.max_column
            except:
                max_col = "?"
            print(f"Max Cols: {max_col}")

            print("Structure Sample (First 10 rows):")
            try:
                # 强制遍历前10行，前5列
                for r in range(1, 11):
                    row_data = []
                    for c in range(1, 6):
                        # 使用 cell() 方法强制读取，不依赖 iter_rows
                        try:
                            cell_val = ws.cell(row=r, column=c).value
                            if cell_val is None:
                                row_data.append("None")
                            else:
                                txt = str(cell_val).strip().replace('\n', ' ')
                                row_data.append(f"'{txt[:15]}...'" if len(txt) > 20 else f"'{txt}'")
                        except:
                            row_data.append("Err")

                    # 只有当这一行不全是 None 时才打印，避免刷屏，但为了调试，至少打印前3行
                    if r <= 3 or any(x != "None" for x in row_data):
                        print(f"  R{r}: {row_data}")
            except Exception as e:
                print(f"  [读取失败: {e}]")
            print("\n")

    except Exception as e:
        print(f"遍历失败: {e}")
    finally:
        if wb: wb.close()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python analyze_excel.py <filepath>")
        sys.exit(1)

    map_excel_structure(sys.argv[1])
