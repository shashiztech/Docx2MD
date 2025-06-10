#!/usr/bin/env python3
"""
Universal Document to Markdown Converter

This converter can handle both .docx (ZIP-based) and .doc (OLE-based) files.
It creates individual directories per document as requested.
"""

import os
import sys
import zipfile
import re
from pathlib import Path
from typing import Dict, List, Tuple
import xml.etree.ElementTree as ET
import logging
import struct

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class DocxToMarkdownConverter:
    """Universal document converter for both .doc and .docx files"""
    
    def __init__(self, input_file: str, output_folder: str = "TargetMDDirectory"):
        self.input_file = input_file
        self.images_extracted = []
        self.headings = []
        doc_name = Path(input_file).stem
        # If output_folder ends with .md, treat as file, else as directory
        output_path = Path(output_folder)
        if output_path.suffix.lower() == ".md":
            self.doc_output_path = output_path.parent
            self.output_markdown_file = output_path
        else:
            self.doc_output_path = output_path / doc_name
            self.output_markdown_file = self.doc_output_path / f"{doc_name}.md"
        self.images_path = self.doc_output_path / "images"
        self.doc_output_path.mkdir(parents=True, exist_ok=True)
        self.images_path.mkdir(parents=True, exist_ok=True)
        
        logger.info(f"Document output directory: {self.doc_output_path}")
        logger.info(f"Images directory: {self.images_path}")
        
        # Determine file type
        self.file_type = self._detect_file_type()
        logger.info(f"Detected file type: {self.file_type}")
    
    def _detect_file_type(self) -> str:
        """Detect if file is .doc or .docx"""
        try:
            # Check if it's a ZIP file (DOCX)
            if zipfile.is_zipfile(self.input_file):
                return "docx"
            
            # Check if it's an OLE file (DOC)
            with open(self.input_file, 'rb') as f:
                header = f.read(8)
                if header[:4] == b'\xd0\xcf\x11\xe0':
                    return "doc"
                    
            return "unknown"
        except Exception as e:
            logger.error(f"Failed to detect file type: {e}")
            return "unknown"
    
    def extract_images_docx(self) -> Dict[str, str]:
        """Extract images from DOCX file"""
        image_mapping = {}
        image_counter = 1
        
        try:
            with zipfile.ZipFile(self.input_file, 'r') as docx_zip:
                # Find all image files
                image_files = [f for f in docx_zip.namelist() 
                             if f.startswith('word/media/') and 
                             any(f.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.svg'])]
                
                for image_file in image_files:
                    try:
                        # Extract image data
                        image_data = docx_zip.read(image_file)
                        
                        # Generate new filename
                        original_name = os.path.basename(image_file)
                        name, ext = os.path.splitext(original_name)
                        new_filename = f"image_{image_counter:03d}_{name}{ext}"
                        
                        # Save image
                        image_path = self.images_path / new_filename
                        with open(image_path, 'wb') as img_file:
                            img_file.write(image_data)
                        
                        # Store mapping
                        image_mapping[original_name] = new_filename
                        image_mapping[image_file] = new_filename
                        
                        self.images_extracted.append(new_filename)
                        image_counter += 1
                        
                        logger.info(f"Extracted image: {new_filename}")
                        
                    except Exception as e:
                        logger.warning(f"Failed to extract image {image_file}: {e}")
                        
        except Exception as e:
            logger.error(f"Failed to extract images from DOCX: {e}")
            
        return image_mapping
    
    def _process_paragraph_text(self, paragraph, namespaces) -> str:
        """Enhanced paragraph processing with better style, bullet/list, and hyperlink preservation"""
        full_text = ""
        is_bullet = False
        is_numbered = False
        bullet_char = "- "
        number_val = None
        # Detect bullet/numbered list
        ppr = paragraph.find('.//w:pPr', namespaces)
        if ppr is not None:
            numpr = ppr.find('.//w:numPr', namespaces)
            if numpr is not None:
                ilvl = numpr.find('.//w:ilvl', namespaces)
                numid = numpr.find('.//w:numId', namespaces)
                if numid is not None:
                    # For simplicity, treat all as bullet unless we can detect actual numbering
                    is_bullet = True
        # Process runs and hyperlinks in order
        for child in paragraph:
            if child.tag.endswith('}hyperlink'):
                # Handle hyperlinks
                link_id = child.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                href = None
                if link_id and hasattr(self, 'hyperlink_mapping'):
                    href = self.hyperlink_mapping.get(link_id)
                link_text = ""
                for run in child.findall('.//w:r', namespaces):
                    run_text = ""
                    for text_elem in run.findall('.//w:t', namespaces):
                        if text_elem.text:
                            run_text += text_elem.text
                    run_props = run.find('.//w:rPr', namespaces)
                    formatted_text = run_text
                    if run_props is not None:
                        if run_props.find('.//w:b', namespaces) is not None:
                            formatted_text = f"**{formatted_text}**"
                        if run_props.find('.//w:i', namespaces) is not None:
                            formatted_text = f"*{formatted_text}*"
                        if run_props.find('.//w:u', namespaces) is not None:
                            if not formatted_text.startswith('**'):
                                formatted_text = f"**{formatted_text}**"
                    link_text += formatted_text
                if href:
                    full_text += f"[{link_text}]({href})"
                else:
                    full_text += link_text
            elif child.tag.endswith('}r'):
                run_text = ""
                for text_elem in child.findall('.//w:t', namespaces):
                    if text_elem.text:
                        run_text += text_elem.text
                if not run_text:
                    continue
                run_props = child.find('.//w:rPr', namespaces)
                formatted_text = run_text
                if run_props is not None:
                    if run_props.find('.//w:b', namespaces) is not None:
                        formatted_text = f"**{formatted_text}**"
                    if run_props.find('.//w:i', namespaces) is not None:
                        formatted_text = f"*{formatted_text}*"
                    if run_props.find('.//w:u', namespaces) is not None:
                        if not formatted_text.startswith('**'):
                            formatted_text = f"**{formatted_text}**"
                full_text += formatted_text
        # Prepend bullet or number if detected
        if is_bullet and full_text.strip():
            return f"- {full_text.strip()}"
        return full_text.strip()

    def _get_hyperlink_mappings(self, docx_zip) -> Dict[str, str]:
        """Get hyperlink mappings from document.xml.rels"""
        hyperlink_mapping = {}
        try:
            if 'word/_rels/document.xml.rels' in docx_zip.namelist():
                rels_xml = docx_zip.read('word/_rels/document.xml.rels').decode('utf-8')
                rels_root = ET.fromstring(rels_xml)
                for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                    rel_id = rel.get('Id')
                    target = rel.get('Target')
                    rel_type = rel.get('Type')
                    if rel_id and target and rel_type and rel_type.endswith('/hyperlink'):
                        hyperlink_mapping[rel_id] = target
        except Exception as e:
            logger.warning(f"Failed to read hyperlink mappings: {e}")
        return hyperlink_mapping

    def extract_text_docx(self) -> str:
        """Enhanced DOCX text extraction with better formatting and hyperlink preservation"""
        text_content = ""
        try:
            with zipfile.ZipFile(self.input_file, 'r') as docx_zip:
                if 'word/document.xml' not in docx_zip.namelist():
                    logger.error("No word/document.xml found in DOCX file")
                    return text_content
                document_xml = docx_zip.read('word/document.xml').decode('utf-8')
                root = ET.fromstring(document_xml)
                namespaces = {
                    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
                    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
                    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                }
                relationship_mapping = self._get_relationship_mappings(docx_zip)
                self.hyperlink_mapping = self._get_hyperlink_mappings(docx_zip)
                body = root.find('.//w:body', namespaces)
                if body is not None:
                    for child in body:
                        if child.tag.endswith('}p'):
                            para_content = self._process_paragraph_with_images(child, namespaces, relationship_mapping)
                            if para_content:
                                text_content += para_content
                        elif child.tag.endswith('}tbl'):
                            table_md = self._convert_table_to_markdown(child, namespaces, relationship_mapping)
                            if table_md:
                                text_content += table_md + "\n\n"
        except Exception as e:
            logger.error(f"Failed to extract text from DOCX: {e}")
        return text_content
    
    def extract_text_doc(self) -> str:
        """Extract text from DOC file using basic parsing"""
        text_content = ""
        
        try:
            # For .doc files, we'll try to extract readable text
            # This is a basic approach that looks for text patterns
            with open(self.input_file, 'rb') as f:
                content = f.read()
                
            # Convert to string and extract readable text
            try:
                # Try different encodings
                for encoding in ['utf-8', 'latin-1', 'cp1252', 'utf-16']:
                    try:
                        text = content.decode(encoding, errors='ignore')
                        break
                    except:
                        continue
                else:
                    text = content.decode('latin-1', errors='ignore')
                
                # Extract readable text using regex
                # Look for sequences of printable characters
                readable_parts = re.findall(r'[A-Za-z0-9\s\.,;:!?\-\(\)\[\]\"\']{10,}', text)
                
                # Filter and clean the text
                cleaned_parts = []
                for part in readable_parts:
                    # Remove excessive whitespace
                    clean_part = re.sub(r'\s+', ' ', part.strip())
                    if len(clean_part) > 15 and not re.match(r'^[0-9\s\.\-]+$', clean_part):
                        cleaned_parts.append(clean_part)
                
                # Join the parts
                if cleaned_parts:
                    text_content = '\n\n'.join(cleaned_parts)
                    
                    # Try to identify headings (lines that are shorter and might be titles)
                    lines = text_content.split('\n')
                    processed_lines = []
                    
                    for line in lines:
                        line = line.strip()
                        if line:
                            # Heuristic: if line is short and doesn't end with punctuation, might be heading
                            if len(line) < 60 and not line.endswith(('.', '!', '?', ',')):
                                # Check if it looks like a heading
                                if len(line.split()) <= 8 and any(word[0].isupper() for word in line.split() if word):
                                    processed_lines.append(f"## {line}")
                                    self.headings.append((2, line))
                                else:
                                    processed_lines.append(line)
                            else:
                                processed_lines.append(line)
                    
                    text_content = '\n\n'.join(processed_lines)
                else:
                    text_content = "Could not extract readable text from this .doc file."
                    
            except Exception as e:
                logger.error(f"Failed to decode .doc file content: {e}")
                text_content = "Error: Could not decode the .doc file content."
                
        except Exception as e:
            logger.error(f"Failed to extract text from DOC: {e}")
            text_content = f"Error extracting from .doc file: {e}"            
        return text_content
    
    def _get_relationship_mappings(self, docx_zip) -> Dict[str, str]:
        """Get relationship mappings from document.xml.rels"""
        relationship_mapping = {}
        
        try:
            if 'word/_rels/document.xml.rels' in docx_zip.namelist():
                rels_xml = docx_zip.read('word/_rels/document.xml.rels').decode('utf-8')
                rels_root = ET.fromstring(rels_xml)
                
                for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                    rel_id = rel.get('Id')
                    target = rel.get('Target')
                    if rel_id and target and target.startswith('media/'):
                        relationship_mapping[rel_id] = target
                        
        except Exception as e:
            logger.warning(f"Failed to read relationship mappings: {e}")
            
        return relationship_mapping
    
    def _process_paragraph_with_images(self, para, namespaces, relationship_mapping) -> str:
        """Process a paragraph and handle any inline images, hyperlinks, and whitespace preservation, with improved heading/subheading detection"""
        parts = []
        is_fully_bold = False
        is_heading = False
        heading_level = 0
        para_text = ""
        bold_text = ""
        # Check if the paragraph is fully bold (all runs bold)
        runs = para.findall('.//w:r', namespaces)
        if runs:
            all_bold = True
            for run in runs:
                run_props = run.find('.//w:rPr', namespaces)
                if not (run_props is not None and run_props.find('.//w:b', namespaces) is not None):
                    all_bold = False
                    break
            if all_bold:
                is_fully_bold = True
        # Build up the paragraph content in order, handling text, hyperlinks, and images
        for child in para:
            if child.tag.endswith('}hyperlink'):
                link_id = child.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                href = None
                if link_id and hasattr(self, 'hyperlink_mapping'):
                    href = self.hyperlink_mapping.get(link_id)
                link_text = ""
                for run in child.findall('.//w:r', namespaces):
                    run_text = ""
                    for text_elem in run.findall('.//w:t', namespaces):
                        if text_elem.text:
                            run_text += text_elem.text
                    run_props = run.find('.//w:rPr', namespaces)
                    formatted_text = run_text
                    if run_props is not None:
                        if run_props.find('.//w:b', namespaces) is not None:
                            formatted_text = f"**{formatted_text}**"
                        if run_props.find('.//w:i', namespaces) is not None:
                            formatted_text = f"*{formatted_text}*"
                        if run_props.find('.//w:u', namespaces) is not None:
                            if not formatted_text.startswith('**'):
                                formatted_text = f"**{formatted_text}**"
                    link_text += formatted_text
                if href:
                    parts.append(f"[{link_text}]({href})")
                else:
                    parts.append(link_text)
            elif child.tag.endswith('}r'):
                run_text = ""
                for text_elem in child.findall('.//w:t', namespaces):
                    if text_elem.text:
                        run_text += text_elem.text
                run_props = child.find('.//w:rPr', namespaces)
                formatted_text = run_text
                is_bold = False
                if run_props is not None:
                    if run_props.find('.//w:b', namespaces) is not None:
                        formatted_text = f"**{formatted_text}**"
                        is_bold = True
                    if run_props.find('.//w:i', namespaces) is not None:
                        formatted_text = f"*{formatted_text}*"
                    if run_props.find('.//w:u', namespaces) is not None:
                        if not formatted_text.startswith('**'):
                            formatted_text = f"**{formatted_text}**"
                # Add images in this run
                drawings = child.findall('.//w:drawing', namespaces)
                if drawings:
                    for drawing in drawings:
                        image_md = self._process_drawing(drawing, namespaces, relationship_mapping)
                        if image_md:
                            if formatted_text:
                                parts.append(formatted_text)
                            parts.append(image_md.strip())
                            formatted_text = ""
                    if formatted_text:
                        parts.append(formatted_text)
                else:
                    if formatted_text:
                        parts.append(formatted_text)
                # For heading/subheading detection
                if is_bold:
                    bold_text += run_text
                para_text += run_text
            elif child.tag.endswith('}drawing'):
                image_md = self._process_drawing(child, namespaces, relationship_mapping)
                if image_md:
                    parts.append(image_md.strip())
        para_text = ''.join(parts).strip()
        # Heuristic: treat fully bold paragraphs as headings
        if is_fully_bold and para_text:
            # If short, treat as H2, else H3
            heading_level = 2 if len(para_text) < 60 else 3
            self.headings.append((heading_level, para_text))
            return f"{'#' * heading_level} {para_text}\n\n"
        # Heuristic: treat bold label + colon as subheading
        if para_text:
            import re
            match = re.match(r'^\*\*(.+?):\*\*\s*(.*)$', para_text)
            if match:
                label = match.group(1).strip()
                rest = match.group(2).strip()
                heading_level = 3
                self.headings.append((heading_level, label))
                if rest:
                    return f"{'#' * heading_level} {label}\n\n{rest}\n\n"
                else:
                    return f"{'#' * heading_level} {label}\n\n"
        # Preserve blank lines and whitespace by checking for empty paragraphs
        if not para_text:
            return "\n"
        # Fallback: use style-based heading detection
        style_elem = para.find('.//w:pStyle', namespaces)
        heading_level = 1
        if style_elem is not None:
            style_val = style_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '')
            if 'Heading' in style_val:
                match = re.search(r'(\d+)', style_val)
                if match:
                    heading_level = min(int(match.group(1)), 6)
                self.headings.append((heading_level, para_text))
                return f"{'#' * heading_level} {para_text}\n\n"
        # Detect if this is a bullet or numbered list
        is_bullet = False
        ppr = para.find('.//w:pPr', namespaces)
        if ppr is not None:
            numpr = ppr.find('.//w:numPr', namespaces)
            if numpr is not None:
                is_bullet = True
        # Format output
        if para_text.startswith('!['):
            # Image only, do not indent
            return para_text + '\n'
        elif is_bullet:
            # List item, do not indent
            return para_text + '\n'
        else:
            # No tab/indent, just return the paragraph
            return f"{para_text}\n\n"
    
    def extract_text(self) -> str:
        """Extract text from the document (DOC or DOCX)"""
        if self.file_type == "docx":
            return self.extract_text_docx()
        elif self.file_type == "doc":
            return self.extract_text_doc()
        return ""
    
    def _process_drawing(self, drawing, namespaces, relationship_mapping) -> str:
        """Process a drawing element and return markdown for any images"""
        try:
            # Look for image references
            blips = drawing.findall('.//a:blip', namespaces)
            
            for blip in blips:
                embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if embed_id and embed_id in relationship_mapping:
                    image_path = relationship_mapping[embed_id]
                    image_filename = os.path.basename(image_path)
                    
                    # Find the corresponding extracted image
                    for extracted_image in self.images_extracted:
                        if image_filename in extracted_image or extracted_image.endswith(f"_{image_filename}"):
                            return f"![{extracted_image}](images/{extracted_image})\n\n"
                    
                    # If not found in extracted images, try to match by original name
                    for extracted_image in self.images_extracted:
                        if image_filename.lower() in extracted_image.lower():
                                                        return f"![{extracted_image}](images/{extracted_image})\n\n"
                            
        except Exception as e:
            logger.warning(f"Failed to process drawing: {e}")
            
        return ""
    
    def _convert_table_to_markdown(self, table_elem, namespaces, relationship_mapping) -> str:
        """Convert a table element to markdown with enhanced formatting"""
        try:
            rows = table_elem.findall('.//w:tr', namespaces)
            if not rows:
                return ""
            table_data = []
            for row in rows:
                cells = row.findall('.//w:tc', namespaces)
                row_data = []
                for cell in cells:
                    paras = cell.findall('.//w:p', namespaces)
                    para_texts = []
                    for para in paras:
                        para_content = self._process_paragraph_with_images(para, namespaces, relationship_mapping)
                        if para_content.strip():
                            para_texts.append(para_content.strip())
                    # Join with double newlines to preserve markdown formatting and paragraph breaks
                    cell_text = '\n\n'.join(para_texts).strip()
                    # If cell_text starts with bold and has a colon, split to next line (handle colon before closing **) and trailing spaces)
                    if cell_text.startswith('**') and ':' in cell_text:
                        # Find the first colon after '**'
                        colon_idx = cell_text.find(':')
                        bold_end = cell_text.find('**', 2)
                        # Check if colon is just before the closing ** (allow spaces)
                        if bold_end != -1 and (colon_idx < bold_end or (colon_idx < bold_end + 3 and cell_text[colon_idx+1:bold_end].strip() == '')):
                            # Move the bolded label (including trailing **) to its own line
                            label = cell_text[:bold_end+2].strip()
                            rest = cell_text[bold_end+2:].strip()
                            cell_text = f"{label}\n{rest}"
                    row_data.append(cell_text)
                if any(cell.strip() for cell in row_data):
                    table_data.append(row_data)
            if not table_data:
                return ""
            return self._format_table_by_type(table_data)
        except Exception as e:
            logger.warning(f"Failed to convert table: {e}")
            return ""
    
    def _format_table_by_type(self, table_data) -> str:
        """Format table based on detected type"""
        if not table_data:
            return ""
        
        num_cols = len(table_data[0])
        num_rows = len(table_data)
        
        # Detect two-column layout with image+text or key-value pairs
        if num_cols == 2 and num_rows >= 1:
            return self._format_two_column_table(table_data)
        
        # Detect complex multi-category tables
        elif num_cols == 3 and num_rows > 3:
            return self._format_multi_category_table(table_data)
        
        # Default markdown table formatting for other cases
        else:
            return self._format_standard_table(table_data)
    
    def _format_standard_table(self, table_data) -> str:
        """Format table as standard markdown table"""
        if not table_data:
            return ""
        
        # Get the maximum number of columns
        max_cols = max(len(row) for row in table_data)
        
        # Pad rows to have the same number of columns
        for row in table_data:
            while len(row) < max_cols:
                row.append("")
        
        # Format as markdown table
        lines = []
        
        # Header row
        if table_data:
            header_row = table_data[0]
            lines.append("| " + " | ".join(cell.replace('\n', ' ').strip() for cell in header_row) + " |")
            lines.append("| " + " | ".join("---" for _ in header_row) + " |")
            
            # Data rows
            for row in table_data[1:]:
                lines.append("| " + " | ".join(cell.replace('\n', ' ').strip() for cell in row) + " |")
        
        return "\n".join(lines) + "\n\n"
    
    def _format_two_column_table(self, table_data) -> str:
        """Enhanced two-column layout formatting with better text and image handling"""
        formatted_sections = []
        for row_idx, row in enumerate(table_data):
            if row_idx == 0:
                continue
            if len(row) >= 2:
                left_content = row[0].strip()
                right_content = row[1].strip()
                # If either cell contains a markdown image, render as a markdown table row
                if ('![' in left_content and not '![' in right_content) or ('![' in right_content and not '![' in left_content):
                    formatted_sections.append(f"| {left_content} | {right_content} |")
                elif left_content and right_content:
                    formatted_sections.append(f"**{left_content}**\n")
                    paragraphs = right_content.split('\n\n')
                    for para in paragraphs:
                        if para.strip():
                            formatted_sections.append(f"{para.strip()}\n")
                    formatted_sections.append("")
        return "\n".join(formatted_sections)
    
    def _format_multi_category_table(self, table_data) -> str:
        """Enhanced multi-category table formatting with better indentation preservation"""
        formatted_sections = []
        current_category = None
        for row_idx, row in enumerate(table_data):
            if len(row) < 3:
                continue
            category, component, description = row[0].strip(), row[1].strip(), row[2].strip()
            # Skip header rows more intelligently
            if (
                category.lower().replace('*', '').strip() in ['category', 'categories'] or 
                component.lower().replace('*', '').strip() in ['component', 'components'] or
                description.lower().replace('*', '').strip() in ['description', 'descriptions']
            ):
                continue
            if not category and not component and not description:
                continue
            if category and category != current_category:
                current_category = category
                formatted_sections.append(f"\n### {category}\n")
            if component or description:
                if component:
                    component_formatted = f"**{component}**"
                    if description:
                        description_lines = description.split('\n')
                        formatted_description = self._format_description_with_indentation(description_lines)
                        formatted_sections.append(f"{component_formatted}: {formatted_description}\n")
                    else:
                        formatted_sections.append(f"{component_formatted}\n")
                elif description:
                    description_lines = description.split('\n')
                    formatted_description = self._format_description_with_indentation(description_lines)
                    formatted_sections.append(f"{formatted_description}\n")
        return "\n".join(formatted_sections) + "\n\n"
    
    def _format_description_with_indentation(self, description_lines) -> str:
        if not description_lines:
            return ""
        formatted_lines = []
        for line in description_lines:
            line = line.strip()
            if not line:
                continue
            if line.startswith(('•', '-', '*')) or (len(line) > 2 and line[1] in '.)'  and line[0].isdigit()):
                formatted_lines.append(f"\n{line}")
            elif line.startswith('    ') or line.startswith('\t'):
                formatted_lines.append(f"\n  {line.lstrip()}")
            else:
                if formatted_lines:
                    formatted_lines.append(f" {line}")
                else:
                    formatted_lines.append(line)
        return "".join(formatted_lines)

    def create_table_of_contents(self) -> str:
        """Create table of contents from headings"""
        if not self.headings:
            return ""
            
        toc = ["## Table of Contents\n"]
        
        for level, heading in self.headings:
            indent = "  " * (level - 1)
            # Create anchor link
            anchor = heading.lower().replace(" ", "-").replace(".", "").replace(",", "")
            anchor = re.sub(r'[^\w\-]', '', anchor)
            toc.append(f"{indent}- [{heading}](#{anchor})")
            
        return "\n".join(toc) + "\n\n"
    
    def convert(self) -> str:
        """Main conversion method"""
        logger.info(f"Starting conversion of {self.input_file}")
        
        # Extract content based on file type
        if self.file_type == "docx":
            # Extract images
            image_mapping = self.extract_images_docx()
            logger.info(f"Extracted {len(self.images_extracted)} images")
            
            # Extract text content
            content = self.extract_text_docx()
        elif self.file_type == "doc":
            # .doc files don't typically have easily extractable images
            logger.info("DOC file detected - image extraction not supported")
            
            # Extract text content
            content = self.extract_text_doc()
        else:
            content = f"Error: Unsupported file type. File '{self.input_file}' is not a recognized document format."
        
        # Add document title
        doc_name = Path(self.input_file).stem
        markdown_content = f"# {doc_name}\n\n"
        
        # Add file type info
        markdown_content += f"*Document Type: {self.file_type.upper()}*\n\n"
        
        # Add table of contents
        if self.headings:
            toc = self.create_table_of_contents()
            markdown_content += toc
          # Add main content
        markdown_content += content
        
        # Note: Images are now inline in the content, no need to add them at the end
        
        return markdown_content
    
    def save_markdown(self, content: str) -> str:
        """Save markdown content to file"""
        output_file = self.output_markdown_file
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(content)
            logger.info(f"Markdown saved to: {output_file}")
            return str(output_file)
        except Exception as e:
            logger.error(f"Failed to save markdown: {e}")
            raise
    
    def generate_report(self) -> str:
        """Generate conversion report"""
        report = f"""# Conversion Report

## Document Information
- **Source File**: {self.input_file}
- **File Type**: {self.file_type.upper()}
- **Output Directory**: {self.doc_output_path}
- **Images Directory**: {self.images_path}

## Conversion Results
- **Total Headings**: {len(self.headings)}
- **Images Extracted**: {len(self.images_extracted)}
- **Output Created**: ✅

## Extracted Images
"""
        
        for i, image in enumerate(self.images_extracted, 1):
            report += f"{i}. {image}\n"
            
        if not self.images_extracted:
            report += "No images found or extracted.\n"
            
        if self.headings:
            report += "\n## Document Structure\n"
            for level, heading in self.headings:
                indent = "  " * (level - 1)
                report += f"{indent}- {heading}\n"
        else:
            report += "\n## Document Structure\nNo headings detected.\n"
                
        return report

def main():
    """Main function"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Convert DOC/DOCX to Markdown with image extraction')
    parser.add_argument('input_file', help='Path to input DOC/DOCX file')
    parser.add_argument('--output', '-o', default='TargetMDDirectory', help='Output folder name')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.input_file):
        print(f"❌ Error: Input file '{args.input_file}' not found.")
        sys.exit(1)
    
    try:
        # Create converter
        converter = DocxToMarkdownConverter(args.input_file, args.output)
        
        # Convert to markdown
        markdown_content = converter.convert()
        
        # Save markdown file
        output_file = converter.save_markdown(markdown_content)
        
        # Generate and save report
        report = converter.generate_report()
        report_file = converter.doc_output_path / "conversion_report.md"
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write(report)
        
        print(f"\n✅ Conversion completed successfully!")
        print(f"📄 Markdown file: {output_file}")
        print(f"🖼️  Images folder: {converter.images_path}")
        print(f"📊 Report: {report_file}")
        print(f"📈 Extracted {len(converter.images_extracted)} images")
        print(f"📋 Found {len(converter.headings)} headings")
        print(f"🗂️  File type: {converter.file_type.upper()}")
        
    except Exception as e:
        print(f"❌ Conversion failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
