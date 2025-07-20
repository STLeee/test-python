import os
import shutil
import sys
import xml.etree.ElementTree as ET
import zipfile

OPENDOCUMENT_NAMESPACES = {
    'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
    'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'draw': 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
    'xlink': 'http://www.w3.org/1999/xlink',
}


def get_style_color(style_element: ET) -> str:
    text_props = style_element.find('style:text-properties', OPENDOCUMENT_NAMESPACES)
    if text_props is not None:
        color = text_props.get(f"{{{OPENDOCUMENT_NAMESPACES['fo']}}}color")
        return color
    return None


def list_red_styles_in_styles_xml(element: ET.Element):
    red_styles = []
    for style in element.findall('.//style:style', OPENDOCUMENT_NAMESPACES):
        style_name = style.get(f"{{{OPENDOCUMENT_NAMESPACES['style']}}}name")
        color = get_style_color(style)
        if color:
            if color == '#ff0000':
                style_name = style_name
                if style_name:
                    red_styles.append(style_name)
        else:
            parent_style_name = style.get(f"{{{OPENDOCUMENT_NAMESPACES['style']}}}parent-style-name")
            print(f"  Parent style: {parent_style_name}")
            if parent_style_name in red_styles:
                red_styles.append(style_name)
    return red_styles


def list_objects_with_red_style_in_content_xml(content_element: ET.Element, red_styles: list[str]):
    red_object_path = []
    for elem in content_element.findall('.//text:p', OPENDOCUMENT_NAMESPACES):
        style_name = elem.get(f"{{{OPENDOCUMENT_NAMESPACES['text']}}}style-name")
        if style_name in red_styles:
            frames = elem.findall('.//draw:frame', OPENDOCUMENT_NAMESPACES)
            for frame in frames:
                object = frame.find('.//draw:object', OPENDOCUMENT_NAMESPACES)
                if object is not None:
                    object_path = object.get(f"{{{OPENDOCUMENT_NAMESPACES['xlink']}}}href")
                    if object_path:
                        red_object_path.append(object_path)
    return red_object_path


def change_formula_color_in_object_content_xml(content_element: ET.Element, namespaces: dict):
    # Find the <semantics> element
    semantics: ET.Element = content_element.find('m:semantics', namespaces)
    if semantics is None:
        raise ValueError("No <semantics> element found in the content XML.")

    # Find the <mfrac> and <annotation> elements
    frac_element: ET.Element = semantics.find('m:mfrac', namespaces)
    annotation_element: ET.Element = semantics.find('m:annotation', namespaces)
    if frac_element is None or annotation_element is None:
        raise ValueError("No <mfrac> or <annotation> element found in the content XML.")

    # Remove <mfrac> from its parent element
    frac_index = list(semantics).index(frac_element)
    semantics.remove(frac_element)

    # Create a new <mstyle> element and set its color attribute
    style_element = ET.Element('mstyle', {'mathcolor': 'red'})

    # Append the <mfrac> element to the new <mstyle> element
    style_element.append(frac_element)

    # Insert the new <mstyle> element back into its original position
    semantics.insert(frac_index, style_element)

    # Modify the <annotation> text content
    original_text = annotation_element.text
    annotation_element.text = f"color red {{{original_text}}}"


def process_odt_document(file_path: str):
    tmp_folder = "tmp_odt"
    if os.path.exists(tmp_folder):
        shutil.rmtree(tmp_folder)
    os.makedirs(tmp_folder, exist_ok=True)

    # unzip the ODT file
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(tmp_folder)

    # list style names with RED color in the styles.xml
    styles_file = os.path.join(tmp_folder, 'styles.xml')
    if not os.path.exists(styles_file):
        print(f"Error: {styles_file} does not exist.")
        return
    root = ET.parse(styles_file).getroot()
    # find all <style:style> elements in the document
    red_styles = list_red_styles_in_styles_xml(root)
    if red_styles:
        print("Styles with RED color:")
        for style in red_styles:
            print(f"  - {style}")
    else:
        print("  No styles with RED color found.")

    # find all objects in text with RED color style in content.xml
    content_file = os.path.join(tmp_folder, 'content.xml')
    if not os.path.exists(content_file):
        print(f"Error: {content_file} does not exist.")
        return
    root = ET.parse(content_file).getroot()
    red_objects = list_objects_with_red_style_in_content_xml(root, red_styles)
    if red_objects:
        print("Objects with RED color style in content:")
        for obj in red_objects:
            print(f"  - {obj}")
    else:
        print("  No objects with RED color style found in content.")

    # Modify the object content.xml to change the color of formulas
    namespace = 'http://www.w3.org/1998/Math/MathML'
    namespaces = {'m': namespace}
    ET.register_namespace('', namespace)
    for obj in red_objects:
        object_content_file = os.path.join(tmp_folder, obj, 'content.xml')
        root = ET.parse(object_content_file).getroot()
        change_formula_color_in_object_content_xml(root, namespaces)
        # Save the modified content.xml back to the file
        root_str = ET.tostring(root, encoding="UTF-8", xml_declaration=True).decode('utf-8')
        print(f"Modified content for object {obj}:\n{root_str}")
        with open(object_content_file, 'w', encoding='utf-8') as f:
            f.write(root_str)
    # Remove the namespace registration
    ET.register_namespace('', '')

    # Repackage the ODT file
    odt_file = file_path.replace('.odt', '_modified.odt')
    with zipfile.ZipFile(odt_file, 'w') as zip_ref:
        for foldername, subfolders, filenames in os.walk(tmp_folder):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                arcname = os.path.relpath(file_path, tmp_folder)
                zip_ref.write(file_path, arcname)


def main():
    odt_file = "test_odt.odt"
    process_odt_document(odt_file)


if __name__ == "__main__":
    sys.exit(main())
