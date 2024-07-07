import csv
from docx import Document
import re
import os
import sys

def dir_content(doc_path, num):

    # 检查文件是否存在且不为空
    if not os.path.exists(doc_path) or os.path.getsize(doc_path) == 0:
        print(f"文件 '{doc_path}' 不存在或为空，程序终止。")
        sys.exit()
    """
    从Word文档中提取标题结构和内容，生成一个树形结构，并将标题及其关系写入CSV文件。
    """
    # 提取标题层级
    def get_heading_level(paragraph):
        match = re.match(r'(\d+(\.\d+)+)', paragraph.text.strip())
        return match.group(1).count('.') if match else None

    document = Document(doc_path)     
    doc_name = doc_path.split('/')[-1]
    
    # 当前段落内容
    current_content = []
    number = num
    # 栈，存储标题和标题级别
    stack = [(number, doc_name, 0, [])]
    
    for paragraph in document.paragraphs:
        level = get_heading_level(paragraph)
        if level is not None:
            current_heading = paragraph.text.strip()
            stack[-1][3].extend(current_content)
            current_content = []
            number += 1
            stack.append((number, current_heading, level, [current_heading]))
        else:
            current_content.append(paragraph.text.strip())
    
    stack[-1][3].extend(current_content)

    # 更新每个节点的内容
    for i in range(0, len(stack) - 1):
        j = i + 1
        while j < len(stack) and stack[i][2] < stack[j][2]:
            stack[i][3].extend(stack[j][3])
            j += 1

    # 写入CSV文件
    with open('title_ner.csv', mode='w', newline='', encoding='utf-8') as ner_file, \
         open('title_relation.csv', mode='w', newline='', encoding='utf-8') as relation_file:
        
        ner_writer = csv.writer(ner_file)
        relation_writer = csv.writer(relation_file)
        
        # 写入CSV文件的表头
        ner_writer.writerow(['Title Number', 'Title', 'Parent Title'])
        relation_writer.writerow(['Title', 'Parent Title'])
        
        # 遍历stack
        for i in range(len(stack)):
            current_number, current_title, current_level, current_content = stack[i]
            
            # 查找父标题
            parent_title = None
            for j in range(i - 1, -1, -1):
                if stack[j][2] < current_level:
                    parent_title = stack[j][1]
                    break
            
            # 将标题及其父标题写入 title_ner.csv
            ner_writer.writerow([current_number, current_title, parent_title if parent_title else 'None'])
            
            # 如果有父标题，将包含关系写入 title_relation.csv
            if parent_title:
                relation_writer.writerow([current_title, parent_title])   
        
    
doc_path = '测试.docx'
output_relation_path = 'title_relation.csv'
output_ner_path = 'title_ner.csv'
dir_content(doc_path, num=0)