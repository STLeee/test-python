from enum import StrEnum
import os
import shutil
import sys
import traceback
import xml.etree.ElementTree as ET
import zipfile

UNZIPPED_STYLE_FILE = 'styles.xml'
UNZIPPED_CONTENT_FILE = 'content.xml'
UNZIPPED_OBJECT_CONTENT_FILE = 'content.xml'

OPENDOCUMENT_NAMESPACES = {
    'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
    'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'draw': 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
    'svg': 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
    'xlink': 'http://www.w3.org/1999/xlink',
    'math': 'http://www.w3.org/1998/Math/MathML',
}


class FormulaColor(StrEnum):
    RED = "#FF0000"
    BLUE = "#0000FF"
    GREEN = "#00FF00"
    BLACK = "#000000"


def get_style_color_from_xml(style_element: ET) -> str:
    text_props = style_element.find('style:text-properties', OPENDOCUMENT_NAMESPACES)
    if text_props is not None:
        color = text_props.get(f"{{{OPENDOCUMENT_NAMESPACES['fo']}}}color")
        return color
    return None


def list_color_styles_in_styles_xml(unzipped_odt_path: str, color: FormulaColor) -> list[str]:
    # Load the styles.xml file
    styles_file = os.path.join(unzipped_odt_path, UNZIPPED_STYLE_FILE)
    if not os.path.exists(styles_file):
        raise FileNotFoundError(f"Styles file {styles_file} does not exist.")
    element = ET.parse(styles_file).getroot()

    color_style_name_list = []
    for style in element.findall('.//style:style', OPENDOCUMENT_NAMESPACES):
        style_name = style.get(f"{{{OPENDOCUMENT_NAMESPACES['style']}}}name")
        style_color = get_style_color_from_xml(style)
        if style_color:
            if style_color.lower() == color.value.lower():
                style_name = style_name
                if style_name:
                    color_style_name_list.append(style_name)
        else:
            # Check if the style has a parent style with the specified color
            parent_style_name = style.get(f"{{{OPENDOCUMENT_NAMESPACES['style']}}}parent-style-name")
            if parent_style_name in color_style_name_list:
                color_style_name_list.append(style_name)
    return color_style_name_list


def list_objects_with_style_in_content_xml(unzipped_odt_path: str, style_name_list: list[str]):
    # Load the content.xml file
    content_file = os.path.join(unzipped_odt_path, UNZIPPED_CONTENT_FILE)
    if not os.path.exists(content_file):
        raise FileNotFoundError(f"Content file {content_file} does not exist.")
    content_element = ET.parse(content_file).getroot()

    # List styles with parent style name in style_name_list
    for style in content_element.findall('.//style:style', OPENDOCUMENT_NAMESPACES):
        style_name = style.get(f"{{{OPENDOCUMENT_NAMESPACES['style']}}}name")
        parent_style_name = style.get(f"{{{OPENDOCUMENT_NAMESPACES['style']}}}parent-style-name")
        if parent_style_name in style_name_list:
            style_name_list.append(style_name)

    # Find all objects in text with style in style_name_list
    object_path_list = []
    object_image_path_list = []
    for elem in content_element.findall('.//text:p', OPENDOCUMENT_NAMESPACES):
        style_name = elem.get(f"{{{OPENDOCUMENT_NAMESPACES['text']}}}style-name")
        if style_name in style_name_list:
            frames = elem.findall('.//draw:frame', OPENDOCUMENT_NAMESPACES)
            for frame in frames:
                object = frame.find('.//draw:object', OPENDOCUMENT_NAMESPACES)
                if object is not None:
                    object_path = object.get(f"{{{OPENDOCUMENT_NAMESPACES['xlink']}}}href")
                    if object_path:
                        object_path_list.append(object_path)
                image = frame.find('.//draw:image', OPENDOCUMENT_NAMESPACES)
                if image is not None:
                    image_path = image.get(f"{{{OPENDOCUMENT_NAMESPACES['xlink']}}}href")
                    if image_path:
                        object_image_path_list.append(image_path)
                    frame.remove(image)

    # Save the modified content.xml back to the file
    content_xml_str = ET.tostring(content_element, encoding="UTF-8", xml_declaration=True).decode('utf-8')
    with open(content_file, 'w', encoding='utf-8') as f:
        f.write(content_xml_str)

    # Remove the images from content.xml
    for image_path in object_image_path_list:
        image_file = os.path.join(unzipped_odt_path, image_path)
        if os.path.exists(image_file):
            os.remove(image_file)

    return object_path_list


def change_formula_color_in_object_content_xml(unzipped_odt_path: str, object_path: str, color: FormulaColor):
    # Load object's content.xml file
    object_content_file = os.path.join(unzipped_odt_path, object_path, UNZIPPED_OBJECT_CONTENT_FILE)
    if not os.path.exists(object_content_file):
        raise FileNotFoundError(f"Content file {object_content_file} does not exist.")
    object_content_element = ET.parse(object_content_file).getroot()

    # Find the <semantics> element
    semantics: ET.Element = object_content_element.find("math:semantics", OPENDOCUMENT_NAMESPACES)
    if semantics is None:
        raise ValueError("No <semantics> element found in the content XML.")
    if len(semantics) == 0:
        raise ValueError("The <semantics> element is empty, no content to modify.")

    # Move all children but <annotation> of <semantics> to a new <mstyle> element with the specified color
    style_element = ET.Element("mstyle", {"mathcolor": color.name.lower()})
    for child in list(semantics):
        if child.tag != f"{{{OPENDOCUMENT_NAMESPACES['math']}}}annotation":
            semantics.remove(child)
            style_element.append(child)
    semantics.insert(0, style_element)

    # Modify the <annotation> text content
    annotation_element: ET.Element = semantics.find("math:annotation", OPENDOCUMENT_NAMESPACES)
    if annotation_element is not None:
        original_text = annotation_element.text
        annotation_element.text = f"color {color.name.lower()} {{{original_text}}}"
        print(f"Annotation: {original_text} -> {annotation_element.text}")
    else:
        print("No <annotation> element found in the content XML.")

    # Save the modified content.xml back to the file
    object_content_xml_str = ET.tostring(object_content_element, encoding="UTF-8", xml_declaration=True).decode("utf-8")
    with open(object_content_file, "w", encoding="utf-8") as f:
        f.write(object_content_xml_str)


def change_formula_color_if_styled_in_xmls(unzipped_odt_path: str, color: FormulaColor):
    # List style names with the specified color in the styles.xml
    color_styles = list_color_styles_in_styles_xml(unzipped_odt_path, color)

    # Find all objects in text with the specified color style in content.xml
    red_objects = list_objects_with_style_in_content_xml(unzipped_odt_path, color_styles)

    # Modify the object content.xml to change the color of formulas
    for obj in red_objects:
        change_formula_color_in_object_content_xml(unzipped_odt_path, obj, color)


def modify_formula_style_in_content_xml(unzipped_odt_path: str, formula_base_pt: int = 4.1):
    # Load the content.xml file
    content_file = os.path.join(unzipped_odt_path, UNZIPPED_CONTENT_FILE)
    if not os.path.exists(content_file):
        raise FileNotFoundError(f"Content file {content_file} does not exist.")
    content_element = ET.parse(content_file).getroot()

    # List and modify styles with parent style name 'Formula'
    formula_style_name_list = []
    for style in content_element.findall('.//style:style', OPENDOCUMENT_NAMESPACES):
        style_name = style.get(f"{{{OPENDOCUMENT_NAMESPACES['style']}}}name")
        parent_style_name = style.get(f"{{{OPENDOCUMENT_NAMESPACES['style']}}}parent-style-name")
        if parent_style_name == "Formula":
            formula_style_name_list.append(style_name)

        # Set vertical-pos to "from-top" and remove vertical-rel
        graphic_properties = style.find('style:graphic-properties', OPENDOCUMENT_NAMESPACES)
        if graphic_properties is not None:
            graphic_properties.set(f"{{{OPENDOCUMENT_NAMESPACES['style']}}}vertical-pos", "from-top")
            graphic_properties.attrib.pop(f"{{{OPENDOCUMENT_NAMESPACES['style']}}}vertical-rel", None)
    if not formula_style_name_list:
        raise ValueError("No styles with parent style name 'Formula' found in content.xml.")

    # Modify the frame with formula style
    modified = False
    for frame in content_element.findall('.//draw:frame', OPENDOCUMENT_NAMESPACES):
        style_name = frame.get(f"{{{OPENDOCUMENT_NAMESPACES['draw']}}}style-name")
        if style_name in formula_style_name_list:
            # Get the height
            height_str = frame.get(f"{{{OPENDOCUMENT_NAMESPACES['svg']}}}height")
            if height_str is None:
                continue
            height_num = float(height_str[:-2])  # Remove 'pt' and convert to int
            # Set frame svg:y to -(height/2 + formula_base_pt)
            frame_y = -(height_num / 2 + formula_base_pt)
            frame.set(f"{{{OPENDOCUMENT_NAMESPACES['svg']}}}y", f"{frame_y}pt")
            modified = True

    if modified:
        # Save the modified content.xml back to the file
        content_xml_str = ET.tostring(content_element, encoding="UTF-8", xml_declaration=True).decode("utf-8")
        with open(content_file, "w", encoding="utf-8") as f:
            f.write(content_xml_str)
    else:
        raise ValueError("No frames with formula style found in content.xml.")


def fix_odt_formula_style(odt_file_path: str, color: FormulaColor):
    tmp_folder = f"{os.path.basename(odt_file_path[:odt_file_path.rfind('.')])}_tmp"
    if os.path.exists(tmp_folder):
        # raise FileExistsError(f"Temporary folder {tmp_folder} already exists. Please remove it before proceeding.")
        shutil.rmtree(tmp_folder)
    os.makedirs(tmp_folder, exist_ok=True)

    # Unzip the ODT file
    with zipfile.ZipFile(odt_file_path, 'r') as zip_ref:
        zip_ref.extractall(tmp_folder)

    # Change the formula color if styled in XMLs
    try:
        change_formula_color_if_styled_in_xmls(tmp_folder, color)
    except ValueError as e:
        print(e)

    # Modify the content.xml to set the formula style
    try:
        modify_formula_style_in_content_xml(tmp_folder)
    except ValueError as e:
        print(e)

    # Repackage the ODT file
    new_odt_file = odt_file_path.replace('.odt', '_modified.odt')
    with zipfile.ZipFile(new_odt_file, 'w') as zip_ref:
        for foldername, subfolders, filenames in os.walk(tmp_folder):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                arcname = os.path.relpath(file_path, tmp_folder)
                zip_ref.write(file_path, arcname)

    # Replace the original file with the modified one
    if os.path.exists(odt_file_path):
        os.remove(odt_file_path)
    shutil.move(new_odt_file, odt_file_path)

    # Clean up the temporary folder
    shutil.rmtree(tmp_folder)


def main():
    odt_file = "test_odt.odt"
    try:
        fix_odt_formula_style(odt_file, FormulaColor.RED)
    except Exception:
        print(f"Error occurred: {traceback.format_exc()}")
        return 1
    print(f"The formula color in {odt_file} has been changed to {FormulaColor.RED.name.lower()}.")


if __name__ == "__main__":
    sys.exit(main())
