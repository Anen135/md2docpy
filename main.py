from docx import Document
import re

def parse_markdown_to_docx(input_md, output_docx):
    doc = Document()
    
    with open(input_md, 'r', encoding='utf-8') as f:
        content = f.read()
    
    lines = content.split('\n')
    
    for line in lines:
        line = line.strip()
        
        if not line and '---' not in line:
            continue
        
        # Обработка горизонтальной линии ---
        if line.startswith('---'):
            doc.add_paragraph().add_run().add_break()
            continue
        
        # Обработка заголовков с форматированием (#### **Bold**)
        if line.startswith('#### '):
            add_formatted_heading(doc, line[5:], level=4)
        elif line.startswith('### '):
            add_formatted_heading(doc, line[4:], level=3)
        elif line.startswith('## '):
            add_formatted_heading(doc, line[3:], level=2)
        elif line.startswith('# '):
            add_formatted_heading(doc, line[2:], level=1)
        
        # Обработка списков и обычного текста (как в предыдущей версии)
        elif re.match(r'^\d+\.\d+\.', line):
            p = doc.add_paragraph(style='List Number 2')
            add_formatted_runs(p, line.split(' ', 1)[1])
        elif re.match(r'^\d+\.', line):
            p = doc.add_paragraph(style='List Number')
            add_formatted_runs(p, line.split(' ', 1)[1])
        elif line.startswith('- ') or line.startswith('* ') or line.startswith('• ') or line.startswith(''):
            p = doc.add_paragraph(style='List Bullet')
            add_formatted_runs(p, line[2:])
        else:
            p = doc.add_paragraph()
            add_formatted_runs(p, line)
    
    doc.save(output_docx)
    print(f"Документ сохранён как: {output_docx}")

def add_formatted_heading(doc, text, level):
    """Добавляет заголовок с поддержкой **жирного** и *курсивного* текста."""
    heading = doc.add_heading(level=level)
    add_formatted_runs(heading, text)

def add_formatted_runs(paragraph, text):
    """Добавляет текст с форматированием **жирный** и *курсив*."""
    parts = re.split(r'(\*\*.*?\*\*|\*.*?\*)', text)
    
    for part in parts:
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            paragraph.add_run(part[2:-2]).bold = True
        elif part.startswith('*') and part.endswith('*'):
            paragraph.add_run(part[1:-1]).italic = True
        else:
            paragraph.add_run(part)

if __name__ == "__main__":
    input_md = "D:\TEMP\mdtodocpy\input.md"
    output_docx = "output.docx"
    parse_markdown_to_docx(input_md, output_docx)