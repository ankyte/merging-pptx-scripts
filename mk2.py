import zipfile
import os
import shutil
import xml.etree.ElementTree as ET
from tempfile import mkdtemp
import re

def merge_pptx_files(pptx1_path, pptx2_path, output_path):
    """
    Merge two PPTX files by working with their underlying XML structure.
    
    Args:
        pptx1_path (str): Path to the first PPTX file
        pptx2_path (str): Path to the second PPTX file
        output_path (str): Path where the merged PPTX file will be saved
    """
    # Create a temporary directory for our work
    temp_dir = mkdtemp()
    
    # Extract the first PPTX (this will be our base)
    base_dir = os.path.join(temp_dir, "base")
    os.makedirs(base_dir)
    
    with zipfile.ZipFile(pptx1_path, 'r') as zip_ref:
        zip_ref.extractall(base_dir)
    
    # Extract the second PPTX
    second_dir = os.path.join(temp_dir, "second")
    os.makedirs(second_dir)
    
    with zipfile.ZipFile(pptx2_path, 'r') as zip_ref:
        zip_ref.extractall(second_dir)
    
    # Read presentation.xml files to get the slide counts and relations
    pres1_xml_path = os.path.join(base_dir, "ppt", "presentation.xml")
    pres2_xml_path = os.path.join(second_dir, "ppt", "presentation.xml")
    
    tree1 = ET.parse(pres1_xml_path)
    root1 = tree1.getroot()
    
    tree2 = ET.parse(pres2_xml_path)
    root2 = tree2.getroot()
    
    # Find the namespace
    ns_match = re.match(r'{(.*)}', root1.tag)
    ns = ns_match.group(1) if ns_match else None
    ns_dict = {'p': ns} if ns else {}
    
    # Find sldIdLst element (contains slide references)
    slide_id_list1 = root1.find('.//p:sldIdLst', ns_dict)
    slide_id_list2 = root2.find('.//p:sldIdLst', ns_dict)
    
    if slide_id_list1 is None or slide_id_list2 is None:
        raise ValueError("Could not find slide ID list in one of the presentations")
    
    # Get the highest slide ID from the first presentation
    max_id = 0
    for slide in slide_id_list1.findall('.//p:sldId', ns_dict):
        id_val = int(slide.get('id', '0'))
        max_id = max(max_id, id_val)
    
    # Get the highest rId from the first presentation's relationship file
    rels_path = os.path.join(base_dir, "ppt", "_rels", "presentation.xml.rels")
    rels_tree = ET.parse(rels_path)
    rels_root = rels_tree.getroot()
    
    max_rel_id = 0
    for rel in rels_root.findall('.//*[@Id]'):
        rid = rel.get('Id', '')
        if rid.startswith('rId'):
            try:
                rid_num = int(rid[3:])
                max_rel_id = max(max_rel_id, rid_num)
            except ValueError:
                continue
    
    # Copy slides from the second presentation
    slides_dir1 = os.path.join(base_dir, "ppt", "slides")
    slides_dir2 = os.path.join(second_dir, "ppt", "slides")
    
    # Ensure the slides directory exists in the base
    os.makedirs(slides_dir1, exist_ok=True)
    
    # Create _rels directory under slides if it doesn't exist
    slides_rels_dir1 = os.path.join(slides_dir1, "_rels")
    os.makedirs(slides_rels_dir1, exist_ok=True)
    
    # Process each slide in the second presentation
    slide_mapping = {}  # To keep track of new IDs for slides
    rel_mapping = {}    # To keep track of new relationship IDs
    
    for slide_id in slide_id_list2.findall('.//p:sldId', ns_dict):
        old_id = slide_id.get('id')
        old_rid = slide_id.get(f'{{{ns}}}id') if ns else slide_id.get('r:id')
        
        # Generate new IDs
        max_id += 1
        max_rel_id += 1
        new_id = str(max_id)
        new_rid = f'rId{max_rel_id}'
        
        slide_mapping[old_id] = new_id
        rel_mapping[old_rid] = new_rid
        
        # Find the slide path from relationships
        rel_file_path = os.path.join(second_dir, "ppt", "_rels", "presentation.xml.rels")
        rel_tree = ET.parse(rel_file_path)
        rel_root = rel_tree.getroot()
        
        slide_path = None
        for rel in rel_root.findall(f".//*[@Id='{old_rid}']"):
            slide_path = rel.get('Target')
            break
        
        if not slide_path:
            continue
        
        # Ensure we have the slide file name
        slide_filename = os.path.basename(slide_path)
        old_slide_path = os.path.join(second_dir, "ppt", slide_path)
        new_slide_path = os.path.join(slides_dir1, slide_filename)
        
        # Copy the slide file
        shutil.copy2(old_slide_path, new_slide_path)
        
        # Copy slide relationships if they exist
        old_slide_rels_dir = os.path.join(os.path.dirname(old_slide_path), "_rels")
        old_slide_rels_file = os.path.join(old_slide_rels_dir, f"{slide_filename}.rels")
        
        if os.path.exists(old_slide_rels_file):
            new_slide_rels_file = os.path.join(slides_rels_dir1, f"{slide_filename}.rels")
            shutil.copy2(old_slide_rels_file, new_slide_rels_file)
        
        # Add to presentation.xml
        new_slide_elem = ET.SubElement(slide_id_list1, f'{{{ns}}}sldId' if ns else 'p:sldId')
        new_slide_elem.set('id', new_id)
        new_slide_elem.set(f'{{{ns}}}id' if ns else 'r:id', new_rid)
        
        # Add to presentation.xml.rels
        new_rel_elem = ET.SubElement(rels_root, 'Relationship')
        new_rel_elem.set('Id', new_rid)
        new_rel_elem.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide')
        new_rel_elem.set('Target', f"slides/{slide_filename}")
    
    # Copy chart data and other media if present
    for folder in ['charts', 'media', 'embeddings', 'theme']:
        src_folder = os.path.join(second_dir, "ppt", folder)
        dst_folder = os.path.join(base_dir, "ppt", folder)
        
        if os.path.exists(src_folder):
            os.makedirs(dst_folder, exist_ok=True)
            
            for item in os.listdir(src_folder):
                src_item = os.path.join(src_folder, item)
                dst_item = os.path.join(dst_folder, item)
                
                if os.path.isfile(src_item):
                    shutil.copy2(src_item, dst_item)
                elif os.path.isdir(src_item):
                    if not os.path.exists(dst_item):
                        shutil.copytree(src_item, dst_item)
    
    # Also copy chart relationship files
    charts_rels_src = os.path.join(second_dir, "ppt", "charts", "_rels")
    charts_rels_dst = os.path.join(base_dir, "ppt", "charts", "_rels")
    
    if os.path.exists(charts_rels_src):
        os.makedirs(charts_rels_dst, exist_ok=True)
        
        for item in os.listdir(charts_rels_src):
            src_item = os.path.join(charts_rels_src, item)
            dst_item = os.path.join(charts_rels_dst, item)
            
            if os.path.isfile(src_item):
                shutil.copy2(src_item, dst_item)
    
    # Copy content types
    content_types_path1 = os.path.join(base_dir, "[Content_Types].xml")
    content_types_path2 = os.path.join(second_dir, "[Content_Types].xml")
    
    ct_tree1 = ET.parse(content_types_path1)
    ct_root1 = ct_tree1.getroot()
    
    ct_tree2 = ET.parse(content_types_path2)
    ct_root2 = ct_tree2.getroot()
    
    # Add missing Override elements
    existing_partnames = set()
    for override in ct_root1.findall(".//*[@PartName]"):
        existing_partnames.add(override.get('PartName'))
    
    for override in ct_root2.findall(".//*[@PartName]"):
        partname = override.get('PartName')
        if partname not in existing_partnames:
            ct_root1.append(override)
    
    # Save modified XML files
    tree1.write(pres1_xml_path, encoding='UTF-8', xml_declaration=True)
    rels_tree.write(rels_path, encoding='UTF-8', xml_declaration=True)
    ct_tree1.write(content_types_path1, encoding='UTF-8', xml_declaration=True)
    
    # Create the new PPTX file
    shutil.make_archive(output_path, 'zip', base_dir)
    
    # Rename zip to pptx
    if os.path.exists(output_path + '.pptx'):
        os.remove(output_path + '.pptx')
    os.rename(output_path + '.zip', output_path + '.pptx')
    
    # Clean up temp directory
    shutil.rmtree(temp_dir)
    
    return output_path + '.pptx'

# Example usage
if __name__ == "__main__":
    merged_file = merge_pptx_files(
        "presentation1.pptx", 
        "presentation2.pptx", 
        "merged_presentation"
    )
    print(f"Created merged presentation: {merged_file}")
