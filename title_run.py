import csv
from docx import Document
import re
import os

# 基于数字开头的标题，提取标题级别
# def get_heading_level(paragraph):
#     match = re.match(r'(\d+(\.\d+)+)', paragraph.text.strip())
#     return match.group(1).count('.')  if match else None

def get_heading_level(paragraph):
    if paragraph.style.name.startswith('Heading'):
        # 提取标题级别（Heading 1, Heading 2, etc.）
        level = int(paragraph.style.name.split(' ')[-1])
        return level
    return None


# 提取标题之间的关系
def extract_stacks(doc_path):
    document = Document(doc_path)
    doc_name = os.path.basename(doc_path)
    
    # 当前段落内容
    current_content = []
    number = 0
    # 栈，存储标题和标题级别
    stack = [(number,doc_name, 0,[])]
    for paragraph in document.paragraphs:
        level = get_heading_level(paragraph)
        print(level)
        if level:
            current_heading = paragraph.text.strip()
            stack[-1][3].extend(current_content)
            current_content = []
            number += 1
            stack.append((number,current_heading, level, [current_heading]))
            
        else:
            current_content.append(paragraph.text.strip())
    stack[-1][3].extend(current_content)
    
    for i in range(0, len(stack)-1):
        j = i + 1
        # 确保 j 不会超出 stack 的范围，并检查标题级别差异
        while j < len(stack) and stack[i][2] - stack[j][2] < 0:
            stack[i][3].extend(stack[j][3])
            j += 1
    return stack

def generate_csv_files(stack):
    # 打开两个CSV文件，分别用于写入标题和关系
    with open('title_ner.csv', mode='w', newline='', encoding='utf-8') as ner_file, \
         open('title_relation.csv', mode='w', newline='', encoding='utf-8') as relation_file:
        
        ner_writer = csv.writer(ner_file)
        relation_writer = csv.writer(relation_file)
        
        # 写入CSV文件的表头
        ner_writer.writerow(['Title Number', 'Title', 'Parent Title', 'Content'])
        relation_writer.writerow(['Title', 'Parent Title', 'Relation'])
        
        # 遍历stack
        for i in range(len(stack)):
            current_number, current_title, current_level, current_content = stack[i]
            current_content = ','.join(current_content)
            # 查找父标题
            parent_title = None
            parent_number = 0
            for j in range(i-1, -1, -1):
                if stack[j][2] < current_level:
                    parent_number = stack[j][0]
                    parent_title = stack[j][1]
                    break
            
            # 将标题及其父标题写入 title_ner.csv
            ner_writer.writerow([current_number, current_title, parent_title,current_content if parent_title else 'None'])
            
            # 如果有父标题，将包含关系写入 title_relation.csv
            if current_number:
              relation_writer.writerow([ parent_number,current_number, '包含'])      
        
    
doc_path = '01.docx'
output_relation_path = 'title_relation.csv'
output_ner_path = 'title_ner.csv'
stack = extract_stacks(doc_path)
generate_csv_files(stack)