import sys

from odf import text, math, teletype
from odf.element import Element
from odf.opendocument import load


def process_odt_document(file_path):
    """
    載入 ODT 檔案，遍歷所有文字，檢查其樣式，並偵測公式。

    Args:
        file_path (str): .odt 檔案的路徑。
    """
    try:
        doc = load(file_path)
    except Exception as e:
        print(f"錯誤：無法載入檔案 {file_path}。")
        print(e)
        return

    print(f"--- 開始處理文件：{file_path} ---")

    # 遍歷所有段落並檢查文字和樣式
    print("\n--- 文字與樣式遍歷 ---")
    all_paragraphs = doc.getElementsByType(text.P)
    if not all_paragraphs:
        print("文件中未找到任何段落。")
    else:
        for i, p in enumerate(all_paragraphs):
            # teletype.extractText() 可以提取段落中的純文字
            paragraph_text = teletype.extractText(p)
            if paragraph_text.strip():  # 僅顯示非空段落
                print(f"\n[段落 {i+1}]")
                print(f"  純文字內容: \"{paragraph_text}\"")

                # 遍歷段落內的子元素以獲取更詳細的樣式資訊
                for child_node in p.childNodes:
                    # 文字節點通常是 unicode 字串或 odf.text.Span 等元素
                    node_text = teletype.extractText(child_node).strip()
                    if not node_text:
                        continue

                    style_name = ""
                    # 元素節點才有 style 屬性
                    if isinstance(child_node, Element):
                        # text:style-name 屬性定義了其樣式
                        style_name = child_node.getAttribute("stylename")

                    if style_name:
                        print(f"  - 文字片段: \"{node_text}\" -> 樣式: '{style_name}'")
                    else:
                        # 如果是段落下的直接文字，它會繼承段落的樣式
                        p_style = p.getAttribute("stylename")
                        print(f"  - 文字片段: \"{node_text}\" -> (繼承段落樣式: '{p_style}')")

    # 偵測文件中的所有公式
    print("\n--- 公式偵測 ---")
    all_formulas = doc.getElementsByType(math.Math)

    if not all_formulas:
        print("文件中未找到任何公式。")
    else:
        print(f"在文件中找到 {len(all_formulas)} 個公式：")
        for i, formula_obj in enumerate(all_formulas):
            # 公式通常以 StarMath 格式儲存在 annotation 中
            annotation = formula_obj.getElementsByType(math.Annotation)
            formula_content = "無法提取公式內容"
            if annotation:
                formula_content = teletype.extractText(annotation[0])

            print(f"  [公式 {i+1}]: {formula_content}")

    print("\n--- 文件處理完畢 ---")


def main():
    odt_file = "test_odt.odt"
    process_odt_document(odt_file)


if __name__ == "__main__":
    sys.exit(main())
