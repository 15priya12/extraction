from docx import Document
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import os
import sys

sys.dont_write_bytecode = True
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'


class GenerateParaRefsForDocx:
    def __init__(self, start_index: int = 100000):
        self.data = []
        self.tables = []
        self.table_counter = 0
        self.para_id_counter = start_index
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }

    def generate_guid(self) -> str:
        """Generate an incremental ID for para_id starting from the configured start_index"""
        current_id = self.para_id_counter
        self.para_id_counter += 1
        return str(current_id)

    def extract_insertions_only(self, paragraph: Paragraph) -> str:
        """Extract paragraph text preserving insertions, skipping deletions,
        and keeping only the visible hyperlink text (not the target URL)."""
        parts = []

        for child in paragraph._element:
            tag = child.tag

            if tag.endswith('}hyperlink'):
                text_runs = child.findall('.//w:t', self.namespaces)
                display_text = "".join([t.text for t in text_runs if t.text])
                if display_text:
                    parts.append(display_text)

            elif tag.endswith('}r'):
                for t in child.findall('.//w:t', self.namespaces):
                    if t.text:
                        parts.append(t.text)

            elif tag.endswith('}ins'):
                for t in child.findall('.//w:t', self.namespaces):
                    if t.text:
                        parts.append(t.text)

            elif tag.endswith('}del'):
                continue

        return "".join(parts).strip()

    def extract_text_from_paragraph(self, paragraph: Paragraph) -> str:
        return self.extract_insertions_only(paragraph)

    def extract_cell_content(self, cell: _Cell) -> str:
        """Extract content from a table cell, handling nested tables and paragraphs"""
        content_parts = []
        for element in cell._element:
            if element.tag.endswith('}p'):
                paragraph = Paragraph(element, cell)
                para_text = self.extract_text_from_paragraph(paragraph)
                if para_text:
                    content_parts.append(para_text)
            elif element.tag.endswith('}tbl'):
                nested_table = Table(element, cell)
                nested_table_md = self.process_table(nested_table)
                if nested_table_md:
                    content_parts.append(f"\n{nested_table_md}\n")

        return " ".join(content_parts).strip()

    def process_table(self, table: Table) -> str:
        """Process a table and convert it to markdown table format, skipping empty tables."""
        markdown_table = []
        rows = table.rows
        if not rows:
            return ""

        header_cells = [self.extract_cell_content(cell) for cell in rows[0].cells]
        if not any(cell.strip() for cell in header_cells):
            return ""

        markdown_table.append("| " + " | ".join(header_cells) + " |")
        markdown_table.append("|" + " --- |" * len(header_cells))

        data_rows_added = False
        for row in rows[1:]:
            row_cells = [self.extract_cell_content(cell) for cell in row.cells]
            if any(cell.strip() for cell in row_cells):
                markdown_table.append("| " + " | ".join(row_cells) + " |")
                data_rows_added = True

        if not data_rows_added and all(not c.strip() for c in header_cells):
            return ""

        return "\n".join(markdown_table) if markdown_table else ""

    def extract_plain_text(self, file_path: str) -> str:
        """Fallback method to extract plain text from DOCX"""
        try:
            doc = Document(file_path)
            full_text = []
            table_index = 0

            for paragraph in doc.paragraphs:
                if paragraph.text.strip():
                    para_text = self.extract_text_from_paragraph(paragraph)
                    if para_text:
                        full_text.append(para_text)

            for table in doc.tables:
                table_index += 1
                table_markdown = self.process_table(table)
                if table_markdown:
                    self.tables.append({
                        'index': table_index,
                        'content': table_markdown
                    })
                    full_text.append(f"SEE TABLE_#{table_index} below")

            result = "\n\n".join(full_text)

            if self.tables:
                result += "\n\n"
                for table_info in self.tables:
                    result += f"\nTABLE_#{table_info['index']}:\n{table_info['content']}\n"

            return result

        except Exception as e:
            print(f"Error in extract_plain_text: {e}")
            return ""

    def process_docx_file(self, file_path: str, start_index: int) -> str:
        """Process a DOCX file and return markdown content including headers or short titles"""
        try:
            self.para_id_counter = start_index
            doc = Document(file_path)
            current_header = ""

            for element in doc.element.body:
                if element.tag.endswith('}p'):
                    paragraph = Paragraph(element, doc)
                    para_text = self.extract_text_from_paragraph(paragraph)
                    if not para_text:
                        continue

                    style_name = getattr(paragraph.style, 'name', '').lower() if paragraph.style else ''
                    word_count = len(para_text.split())

                    # Determine if this paragraph should be treated as a header
                    if "heading" in style_name or word_count <= 6:
                        current_header = para_text.strip()
                        continue  # skip adding this as normal content

                    # Normal content
                    self.data.append({
                        'para_id': self.generate_guid(),
                        'header': current_header,
                        'para_text': para_text
                    })

                elif element.tag.endswith('}tbl'):
                    self.table_counter += 1
                    table = Table(element, doc)
                    table_markdown = self.process_table(table)

                    if table_markdown:
                        self.tables.append({
                            'index': self.table_counter,
                            'content': table_markdown
                        })
                        self.data.append({
                            'para_id': self.generate_guid(),
                            'header': current_header,
                            'para_text': f"SEE TABLE_{self.table_counter} below"
                        })

            if self.data:
                return self.generate_markdown_table_with_header()
            else:
                return self.extract_plain_text(file_path)

        except Exception as e:
            print(f"Error processing DOCX: {e}")
            return self.extract_plain_text(file_path)

    def generate_markdown_table_with_header(self) -> str:
        """Generate markdown table with headers"""
        if not self.data:
            return ""

        lines = ["| para_id | header | para_text |", "|---------|---------|-----------|"]

        for item in self.data:
            header = item.get('header', '')
            escaped_text = item['para_text'].replace('|', '\\|').replace('\n', '<br>')
            escaped_header = header.replace('|', '\\|').replace('\n', '<br>')
            lines.append(f"| {item['para_id']} | {escaped_header} | {escaped_text} |")

        result = "\n".join(lines)

        if self.tables:
            result += "\n\n"
            for table_info in self.tables:
                result += f"\nTABLE_{table_info['index']}:\n{table_info['content']}\n"

        return result

    def clear_data(self):
        """Clear the stored data"""
        self.data = []
        self.tables = []
        self.table_counter = 0


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Extract DOCX content with paragraph references and inferred headers.")
    parser.add_argument("input_file", help="Path to the DOCX file to process.")
    parser.add_argument("--start_index", type=int, default=100000, help="Starting index for paragraph IDs.")
    args = parser.parse_args()

    processor = GenerateParaRefsForDocx(start_index=args.start_index)
    markdown_output = processor.process_docx_file(args.input_file, args.start_index)

    output_file = os.path.splitext(args.input_file)[0] + "_output.md"
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(markdown_output)

    print(f"✅ Markdown file generated: {output_file}")
