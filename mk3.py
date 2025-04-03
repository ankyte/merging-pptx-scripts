import zipfile
import os
import shutil
import xml.etree.ElementTree as ET
from tempfile import mkdtemp
import re
import uuid

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
    
    # Read presentation.xml files
    pres1_xml_path = os.path.join(base_dir, "ppt", "presentation.xml")
    pres2_xml_path = os.path.join(second_dir, "ppt", "presentation.xml")
    
    tree1 = ET.parse(pres1_xml_path)
    root1 = tree1.getroot()
    
    tree2 = ET.parse(pres2_xml_path)
    root2 = tree2.getroot()
    
    # Find namespace
    ns_match = re.match(r'{(.*)}', root1.tag)
    ns = ns_match.group(1) if ns_match else None
    ns_dict = {'p': ns, 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'} if ns else {}
    
    # Register namespaces for proper XML output
    for prefix, uri in ns_dict.items():
        ET.register_namespace(prefix, uri)
    
    # Register default namespace for OOXML content types
    ET.register_namespace('', 'http://schemas.openxmlformats.org/package/2006/content-types')
    
    # Register relationships namespace
    ET.register_namespace('r', 'http://schemas.openxmlformats.org/package/2006/relationships')
    
    # Find sldIdLst element (contains slide references)
    slide_id_list1 = root1.find('.//p:sldIdLst', ns_dict)
    slide_id_list2 = root2.find('.//p:sldIdLst', ns_dict)
    
    if slide_id_list1 is None or slide_id_list2 is None:
        if slide_id_list1 is None:
            print("Could not find slide ID list in first presentation")
        if slide_id_list2 is None:
            print("Could not find slide ID list in second presentation")
        raise ValueError("Could not find slide ID list in one of the presentations")
    
    # Get the highest slide ID from the first presentation
    max_id = 0
    for slide in slide_id_list1.findall('./p:sldId', ns_dict):
        id_val = int(slide.get('id', '0'))
        max_id = max(max_id, id_val)
    
    # Get relationships from the first presentation
    rels_path1 = os.path.join(base_dir, "ppt", "_rels", "presentation.xml.rels")
    rels_tree1 = ET.parse(rels_path1)
    rels_root1 = rels_tree1.getroot()
    
    # Get relationships from the second presentation
    rels_path2 = os.path.join(second_dir, "ppt", "_rels", "presentation.xml.rels")
    rels_tree2 = ET.parse(rels_path2)
    rels_root2 = rels_tree2.getroot()
    
    # Find highest rId number in first presentation
    max_rel_id = 0
    for rel in rels_root1.findall('.//*[@Id]'):
        rid = rel.get('Id', '')
        if rid.startswith('rId'):
            try:
                rid_num = int(rid[3:])
                max_rel_id = max(max_rel_id, rid_num)
            except ValueError:
                continue
    
    # Map of old rIds to new rIds for the second presentation
    rid_mapping = {}
    
    # Step 1: Copy all slide content from second presentation
    # This includes slides, slideMasters, slideLayouts, and their relationships
    for folder_name in ['slides', 'slideMasters', 'slideLayouts', 'charts', 'media', 'embeddings', 'theme', 'diagrams']:
        src_folder = os.path.join(second_dir, "ppt", folder_name)
        dst_folder = os.path.join(base_dir, "ppt", folder_name)
        
        if os.path.exists(src_folder):
            os.makedirs(dst_folder, exist_ok=True)
            
            # Copy files
            for item in os.listdir(src_folder):
                src_item = os.path.join(src_folder, item)
                dst_item = os.path.join(dst_folder, item)
                
                # Skip _rels folders for now, we'll handle them separately
                if item == "_rels":
                    continue
                
                if os.path.isfile(src_item) and not os.path.exists(dst_item):
                    shutil.copy2(src_item, dst_item)
                elif os.path.isdir(src_item) and not os.path.exists(dst_item):
                    shutil.copytree(src_item, dst_item)
            
            # Now handle relationship folders
            src_rels = os.path.join(src_folder, "_rels")
            dst_rels = os.path.join(dst_folder, "_rels")
            
            if os.path.exists(src_rels):
                os.makedirs(dst_rels, exist_ok=True)
                
                for item in os.listdir(src_rels):
                    src_item = os.path.join(src_rels, item)
                    dst_item = os.path.join(dst_rels, item)
                    
                    if os.path.isfile(src_item) and not os.path.exists(dst_item):
                        shutil.copy2(src_item, dst_item)
    
    # Step 2: Process content types to include all slide types
    content_types_path1 = os.path.join(base_dir, "[Content_Types].xml")
    content_types_path2 = os.path.join(second_dir, "[Content_Types].xml")
    
    ct_tree1 = ET.parse(content_types_path1)
    ct_root1 = ct_tree1.getroot()
    
    ct_tree2 = ET.parse(content_types_path2)
    ct_root2 = ct_tree2.getroot()
    
    # Add content types from second presentation
    existing_partnames = set()
    for override in ct_root1.findall(".//*[@PartName]"):
        existing_partnames.add(override.get('PartName'))
    
    for override in ct_root2.findall(".//*[@PartName]"):
        partname = override.get('PartName')
        if partname not in existing_partnames:
            ct_root1.append(override)
    
    # Step 3: Add slides from second presentation to first presentation
    for slide_id in slide_id_list2.findall('./p:sldId', ns_dict):
        old_id = slide_id.get('id')
        r_id_attr = f'{{{ns_dict["r"]}}}id' if 'r' in ns_dict else 'r:id'
        old_rid = slide_id.get(r_id_attr)
        
        # Generate new IDs
        max_id += 1
        max_rel_id += 1
        new_id = str(max_id)
        new_rid = f'rId{max_rel_id}'
        
        # Store mapping
        rid_mapping[old_rid] = new_rid
        
        # Find slide path from relationships
        slide_target = None
        for rel in rels_root2.findall(f".//*[@Id='{old_rid}']"):
            slide_target = rel.get('Target')
            break
        
        if not slide_target:
            print(f"Could not find relationship for slide {old_id}")
            continue
        
        # Create new slide element in first presentation
        new_slide_elem = ET.SubElement(slide_id_list1, f'{{{ns}}}sldId' if ns else 'p:sldId')
        new_slide_elem.set('id', new_id)
        new_slide_elem.set(r_id_attr, new_rid)
        
        # Add relationship to presentation.xml.rels
        rel_elem = ET.SubElement(rels_root1, 'Relationship')
        rel_elem.set('Id', new_rid)
        rel_elem.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide')
        rel_elem.set('Target', slide_target)
    
    # Step 4: Add all other relationships from second presentation to first
    # (except slide relationships which we've already handled)
    for rel in rels_root2.findall('.//*[@Id]'):
        rid = rel.get('Id')
        rel_type = rel.get('Type')
        target = rel.get('Target')
        
        # Skip slide relationships (already handled)
        if rel_type == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide':
            continue
        
        # Check if this relationship target already exists
        target_exists = False
        for existing_rel in rels_root1.findall('.//*[@Target]'):
            if existing_rel.get('Target') == target and existing_rel.get('Type') == rel_type:
                target_exists = True
                break
        
        if not target_exists:
            max_rel_id += 1
            new_rid = f'rId{max_rel_id}'
            rid_mapping[rid] = new_rid
            
            # Add new relationship
            new_rel = ET.SubElement(rels_root1, 'Relationship')
            new_rel.set('Id', new_rid)
            new_rel.set('Type', rel_type)
            new_rel.set('Target', target)
            
            # If this is an external relationship (target with http://)
            if rel.get('TargetMode') == 'External':
                new_rel.set('TargetMode', 'External')
    
    # Save modified XML files
    tree1.write(pres1_xml_path, encoding='UTF-8', xml_declaration=True)
    rels_tree1.write(rels_path1, encoding='UTF-8', xml_declaration=True)
    ct_tree1.write(content_types_path1, encoding='UTF-8', xml_declaration=True)
    
    # Create new PPTX file
    output_pptx = output_path
    if not output_pptx.endswith('.pptx'):
        output_pptx += '.pptx'
    
    if os.path.exists(output_pptx):
        os.remove(output_pptx)
    
    # Create the PPTX file (zip the directory)
    with zipfile.ZipFile(output_pptx, 'w') as zipf:
        for root, dirs, files in os.walk(base_dir):
            for file in files:
                file_path = os.path.join(root, file)
                zipf.write(file_path, os.path.relpath(file_path, base_dir))
    
    # Clean up temp directory
    shutil.rmtree(temp_dir)
    
    return output_pptx

# Example usage
if __name__ == "__main__":
    merged_file = merge_pptx_files(
        "presentation1.pptx", 
        "presentation2.pptx", 
        "merged_presentation.pptx"
    )
    print(f"Created merged presentation: {merged_file}")
