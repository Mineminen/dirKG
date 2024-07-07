from docx import Document

def get_heading_level(paragraph):
    # 检查段落样式是否为标题
    if paragraph.style.name.startswith('Heading'):
        # 提取标题级别（Heading 1, Heading 2, etc.）
        level = int(paragraph.style.name.split(' ')[-1])
        return level
    return None

# 读取Word文档
doc = Document('01.docx')

# 遍历所有段落
for para in doc.paragraphs:
    level = get_heading_level(para)
    if level is not None:
        print(f'Title: {para.text}, Level: {level}')
    else:
        print(f'Paragraph: {para.text},level: {level}')


