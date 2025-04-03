import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
import re
import uuid

def merge_presentations(base_ppt_path, add_ppt_path, output_path):
    """
    Merge two PowerPoint presentations using XML manipulation.
    
    Args:
        base_ppt_path (str): Path to the first presentation (base presentation)
        add_ppt_path (str): Path to the second presentation to be added
        output_path (str): Path where the merged presentation will be saved
    """
    # Create temporary directories
    temp_dir = "temp_ppt_merge"
    base_dir = os.path.join(temp_dir, "base")
    add_dir = os.path.join(temp_dir, "add")
    
    try:
        # Create directories if they don't exist
        for dir_path in [temp_dir, base_dir, add_dir]:
            if not os.path.exists(dir_path):
                os.makedirs(dir_path)
        
        # Extract both presentations
        extract_pptx(base_ppt_path, base_dir)
        extract_pptx(add_ppt_path, add_dir)
        
        # Get slide count from base presentation
        presentation_xml_path = os.path.join(base_dir, "ppt", "presentation.xml")
        tree = ET.parse(presentation_xml_path)
        root = tree.getroot()
        
        # Find namespace
        ns = get_namespace(root)
        # Create namespace map for finding elements
        nsmap = {'p': ns}
        
        # Find slide references
        slides_element = root.find('.//p:sldIdLst', nsmap)
        if slides_element is None:
            slides_element = ET.SubElement(root.find('.//p:presentation', nsmap), f'{{{ns}}}sldIdLst')
        
        # Get highest slide ID from base presentation
        slide_ids = [int(slide_id.get('id')) for slide_id in slides_element.findall('./p:sldId', nsmap)]
        max_slide_id = max(slide_ids) if slide_ids else 255
        
        # Get current relationships
        rels_path = os.path.join(base_dir, "ppt", "_rels", "presentation.xml.rels")
        rels_tree = ET.parse(rels_path)
        rels_root = rels_tree.getroot()
        
        # Find highest relationship ID
        rel_ids = [int(rel.get('Id').replace('rId', '')) for rel in rels_root.findall('.//*[@Id]')]
        max_rel_id = max(rel_ids) if rel_ids else 0
        
        # Track new media and relationships
        added_media = {}
        slide_rel_map = {}
        
        # Process each slide in the second presentation
        add_pres_xml_path = os.path.join(add_dir, "ppt", "presentation.xml")
        add_tree = ET.parse(add_pres_xml_path)
        add_root = add_tree.getroot()
        
        # Find slide references in the presentation to be added
        add_slides_element = add_root.find('.//p:sldIdLst', {'p': get_namespace(add_root)})
        if add_slides_element is None:
            print("No slides found in the presentation to be added.")
            return
        
        # Process each slide in the second presentation
        for slide_id_elem in add_slides_element.findall('./p:sldId', {'p': get_namespace(add_root)}):
            # Get the relationship ID for this slide
            rid = slide_id_elem.get(f'{{{get_namespace(add_root, "r")}}}id')
            
            # Find the corresponding relationship
            add_rels_path = os.path.join(add_dir, "ppt", "_rels", "presentation.xml.rels")
            add_rels_tree = ET.parse(add_rels_path)
            add_rels_root = add_rels_tree.getroot()
            
            slide_rel = None
            for rel in add_rels_root.findall('.//*[@Id]'):
                if rel.get('Id') == rid:
                    slide_rel = rel
                    break
            
            if slide_rel is None:
                continue
            
            # Get the slide path
            slide_path = slide_rel.get('Target')
            full_add_slide_path = os.path.join(add_dir, "ppt", slide_path)
            
            if not os.path.exists(full_add_slide_path):
                continue
            
            # Create new slide IDs
            max_slide_id += 1
            max_rel_id += 1
            new_slide_id = str(max_slide_id)
            new_rel_id = f'rId{max_rel_id}'
            
            # Copy slide file
            slide_filename = os.path.basename(slide_path)
            new_slide_path = f"slides/slide{max_slide_id}.xml"
            full_new_slide_path = os.path.join(base_dir, "ppt", new_slide_path)
            
            # Create slides directory if it doesn't exist
            slides_dir = os.path.join(base_dir, "ppt", "slides")
            if not os.path.exists(slides_dir):
                os.makedirs(slides_dir)
            
            # Copy slide XML
            shutil.copy2(full_add_slide_path, full_new_slide_path)
            
            # Add slide reference to presentation.xml
            new_slide_elem = ET.SubElement(slides_element, f'{{{ns}}}sldId')
            new_slide_elem.set('id', new_slide_id)
            new_slide_elem.set(f'{{{get_namespace(root, "r")}}}id', new_rel_id)
            
            # Add relationship to presentation.xml.rels
            new_rel_elem = ET.SubElement(rels_root, f'{{{get_namespace(rels_root)}}}Relationship')
            new_rel_elem.set('Id', new_rel_id)
            new_rel_elem.set('Type', "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
            new_rel_elem.set('Target', new_slide_path)
            
            # Process slide's relationships (images, charts, etc.)
            process_slide_relationships(base_dir, add_dir, slide_path, new_slide_path, added_media)
            
        # Save modified files
        tree.write(presentation_xml_path, encoding='UTF-8', xml_declaration=True)
        rels_tree.write(rels_path, encoding='UTF-8', xml_declaration=True)
        
        # Update content types
        update_content_types(base_dir)
        
        # Create the merged presentation
        create_pptx(base_dir, output_path)
        
        print(f"Successfully merged presentations into {output_path}")
        
    except Exception as e:
        print(f"Error merging presentations: {e}")
    finally:
        # Clean up temporary files
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

def extract_pptx(pptx_path, extract_dir):
    """Extract a .pptx file to the specified directory."""
    with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

def create_pptx(source_dir, output_path):
    """Create a .pptx file from the specified directory."""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for root, _, files in os.walk(source_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, source_dir)
                zip_ref.write(file_path, arcname)

def get_namespace(element, prefix='p'):
    """Extract namespace from an XML element."""
    m = re.match(r'\{(.*)\}', element.tag)
    if m:
        namespace = m.group(1)
        if prefix == 'p':
            return namespace
        elif prefix == 'r':
            return namespace.replace('schemas.openxmlformats.org/presentationml', 
                                   'schemas.openxmlformats.org/officeDocument')
    return ""

def process_slide_relationships(base_dir, add_dir, slide_path, new_slide_path, added_media):
    """Process relationships of a slide, copying media files as needed."""
    # Get slide relationships
    slide_rel_dir = os.path.join(add_dir, "ppt", os.path.dirname(slide_path), "_rels")
    slide_rel_file = os.path.basename(slide_path) + ".rels"
    slide_rel_path = os.path.join(slide_rel_dir, slide_rel_file)
    
    # Create target relationships directory if needed
    target_rel_dir = os.path.join(base_dir, "ppt", os.path.dirname(new_slide_path), "_rels")
    if not os.path.exists(target_rel_dir):
        os.makedirs(target_rel_dir)
    
    target_rel_path = os.path.join(target_rel_dir, os.path.basename(new_slide_path) + ".rels")
    
    if not os.path.exists(slide_rel_path):
        # No relationships to process
        return
    
    # Parse relationships
    rels_tree = ET.parse(slide_rel_path)
    rels_root = rels_tree.getroot()
    
    # Process each relationship
    for rel in rels_root.findall('.//*[@Target]'):
        target = rel.get('Target')
        rel_type = rel.get('Type')
        
        # Handle media files (images, audio, video)
        if any(type_id in rel_type for type_id in ['image', 'audio', 'video']):
            # Check if target is a local path
            if not target.startswith('http'):
                # Get media file path
                media_path = os.path.normpath(os.path.join(os.path.dirname(slide_path), target))
                
                # Compute absolute paths
                source_media_path = os.path.join(add_dir, "ppt", media_path)
                
                # Generate a unique path for the media file
                if media_path not in added_media:
                    # Extract extension
                    _, ext = os.path.splitext(media_path)
                    new_filename = f"media/image{len(added_media) + 1}{ext}"
                    added_media[media_path] = new_filename
                    
                    # Ensure media directory exists
                    media_dir = os.path.join(base_dir, "ppt", "media")
                    if not os.path.exists(media_dir):
                        os.makedirs(media_dir)
                    
                    # Copy media file
                    target_media_path = os.path.join(base_dir, "ppt", new_filename)
                    if os.path.exists(source_media_path):
                        shutil.copy2(source_media_path, target_media_path)
                
                # Update relationship target
                rel.set('Target', f"../{added_media[media_path]}")
        
        # Handle charts
        elif 'chart' in rel_type:
            # Get chart path
            chart_path = os.path.normpath(os.path.join(os.path.dirname(slide_path), target))
            
            # Compute absolute paths
            source_chart_path = os.path.join(add_dir, "ppt", chart_path)
            
            # Generate a unique ID for the chart
            chart_id = f"chart{uuid.uuid4().hex[:8]}"
            new_chart_path = f"charts/{chart_id}.xml"
            
            # Ensure charts directory exists
            charts_dir = os.path.join(base_dir, "ppt", "charts")
            if not os.path.exists(charts_dir):
                os.makedirs(charts_dir)
            
            # Copy chart file
            target_chart_path = os.path.join(base_dir, "ppt", new_chart_path)
            if os.path.exists(source_chart_path):
                shutil.copy2(source_chart_path, target_chart_path)
                
                # Process chart relationships
                chart_rel_dir = os.path.join(add_dir, "ppt", "charts", "_rels")
                chart_rel_file = os.path.basename(chart_path) + ".rels"
                chart_rel_path = os.path.join(chart_rel_dir, chart_rel_file)
                
                # Create target relationships directory if needed
                target_chart_rel_dir = os.path.join(base_dir, "ppt", "charts", "_rels")
                if not os.path.exists(target_chart_rel_dir):
                    os.makedirs(target_chart_rel_dir)
                
                target_chart_rel_path = os.path.join(target_chart_rel_dir, chart_id + ".xml.rels")
                
                if os.path.exists(chart_rel_path):
                    # Copy and update chart relationships
                    process_chart_relationships(chart_rel_path, target_chart_rel_path, base_dir, add_dir)
            
            # Update relationship target
            rel.set('Target', f"../{new_chart_path}")
    
    # Save the updated relationships
    rels_tree.write(target_rel_path, encoding='UTF-8', xml_declaration=True)

def process_chart_relationships(chart_rel_path, target_chart_rel_path, base_dir, add_dir):
    """Process relationships of a chart, copying related files as needed."""
    if not os.path.exists(chart_rel_path):
        return
    
    # Parse relationships
    rels_tree = ET.parse(chart_rel_path)
    rels_root = rels_tree.getroot()
    
    # Process each relationship
    for rel in rels_root.findall('.//*[@Target]'):
        target = rel.get('Target')
        rel_type = rel.get('Type')
        
        # Handle embedded Excel data
        if 'package' in rel_type and '.xlsx' in target:
            # Get Excel path
            excel_path = target
            if excel_path.startswith('../'):
                excel_path = excel_path[3:]  # Remove leading '../'
            
            # Compute absolute paths
            source_excel_path = os.path.join(add_dir, excel_path)
            
            # Generate a unique ID for the Excel file
            excel_id = f"embeddings/Microsoft_Excel_Sheet{uuid.uuid4().hex[:8]}.xlsx"
            
            # Ensure embeddings directory exists
            embed_dir = os.path.join(base_dir, "embeddings")
            if not os.path.exists(embed_dir):
                os.makedirs(embed_dir)
            
            # Copy Excel file
            target_excel_path = os.path.join(base_dir, excel_id)
            if os.path.exists(source_excel_path):
                shutil.copy2(source_excel_path, target_excel_path)
            
            # Update relationship target
            rel.set('Target', f"../{excel_id}")
    
    # Save the updated relationships
    rels_tree.write(target_chart_rel_path, encoding='UTF-8', xml_declaration=True)

def update_content_types(base_dir):
    """Update the [Content_Types].xml file to include all content types."""
    content_types_path = os.path.join(base_dir, "[Content_Types].xml")
    
    if not os.path.exists(content_types_path):
        return
    
    # Parse content types
    tree = ET.parse(content_types_path)
    root = tree.getroot()
    
    # Ensure all needed content types are present
    content_types = {
        "slides": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        "slideMaster": "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
        "slideLayout": "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
        "chart": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
        "drawing": "application/vnd.openxmlformats-officedocument.drawing+xml",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }
    
    extensions = {
        ".jpeg": "image/jpeg",
        ".jpg": "image/jpeg",
        ".png": "image/png",
        ".gif": "image/gif",
        ".wmf": "image/x-wmf",
        ".mp3": "audio/mp3",
        ".mp4": "video/mp4",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }
    
    # Check for existing defaults and add missing ones
    existing_defaults = {ext.get('Extension'): True for ext in root.findall(".//{*}Default")}
    
    for ext, mime in extensions.items():
        ext_clean = ext[1:]  # Remove the leading dot
        if ext_clean not in existing_defaults:
            default = ET.SubElement(root, "{http://schemas.openxmlformats.org/package/2006/content-types}Default")
            default.set('Extension', ext_clean)
            default.set('ContentType', mime)
    
    # Check for existing overrides and add missing ones
    existing_overrides = {}
    for override in root.findall(".//{*}Override"):
        part_name = override.get('PartName')
        if part_name:
            existing_overrides[part_name] = True
    
    # Add overrides for slides
    slide_pattern = re.compile(r'/ppt/slides/slide(\d+)\.xml')
    for slide_file in os.listdir(os.path.join(base_dir, "ppt", "slides")):
        if slide_file.endswith('.xml'):
            part_name = f"/ppt/slides/{slide_file}"
            if part_name not in existing_overrides:
                override = ET.SubElement(root, "{http://schemas.openxmlformats.org/package/2006/content-types}Override")
                override.set('PartName', part_name)
                override.set('ContentType', content_types["slides"])
    
    # Add overrides for charts
    if os.path.exists(os.path.join(base_dir, "ppt", "charts")):
        for chart_file in os.listdir(os.path.join(base_dir, "ppt", "charts")):
            if chart_file.endswith('.xml'):
                part_name = f"/ppt/charts/{chart_file}"
                if part_name not in existing_overrides:
                    override = ET.SubElement(root, "{http://schemas.openxmlformats.org/package/2006/content-types}Override")
                    override.set('PartName', part_name)
                    override.set('ContentType', content_types["chart"])
    
    # Save the updated content types
    tree.write(content_types_path, encoding='UTF-8', xml_declaration=True)

# Example usage
if __name__ == "__main__":
    merge_presentations("base_presentation.pptx", "additional_presentation.pptx", "merged_output.pptx")
