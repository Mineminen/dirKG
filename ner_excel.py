import os
import pandas as pd
from datetime import datetime
import time
from omegaconf import DictConfig
import hydra
from hydra import utils
import re
from docx import Document
import sys


# 获取文件或文件夹时间属性的函数
def get_attributes(path,root_dir):
    try:
        stat = os.stat(path)
        relative_path = os.path.relpath(path, os.path.dirname(root_dir))
        # parent_dir = '' if path == root_dir else os.path.dirname(relative_path)
        attributes = {
                'Name': os.path.basename(path),
                'Path': path,
                'Creation Time':  time.ctime(stat.st_ctime),
                # 'Modification Time': datetime.fromtimestamp(stat.st_mtime),
                # 'Access Time': datetime.fromtimestamp(stat.st_atime),
                'Label': 'Directory' if os.path.isdir(path) else 'File',
                'Title': None,
                'Content': None,
                # 'parent_title': None,   
            }
        if os.path.isfile(path):
            if path.endswith('.txt') or path.endswith('.py'):
                with open(path, 'r', encoding='utf-8') as file:
                    attributes['Content'] = file.read()
        return attributes
        
    except Exception as e:
        print(f"Error processing {path}: {e}")
        return {'Path': os.path.basename(path), 'Creation Time': None, 'Type': None}

def dir_content(doc_path, root_dir):
    all_attributes = []
 
    # 提取标题层级
    def get_heading_level(paragraph):
        if paragraph.style.name.startswith('Heading'):
            # 提取标题级别（Heading 1, Heading 2, etc.）
            level = int(paragraph.style.name.split(' ')[-1])
            return level
        return None

    document = Document(doc_path)     
    doc_name = os.path.basename(doc_path)
    
    # 当前段落内容
    current_content = []
    number = 0
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


    # 遍历stack
    for i in range(len(stack)):
        current_number, current_title, current_level, current_content = stack[i]
        attribute=get_attributes(doc_path,root_dir)
        current_content = ','.join(current_content).strip('"')
        
        # 查找父标题
        parent_title = None
        for j in range(i - 1, -1, -1):
            if stack[j][2] < current_level:
                parent_title = stack[j][1]
                break
        attribute['Content']=current_content
        attribute['Title']=current_title
        # attribute['parent_title']=parent_title
        all_attributes.append(attribute) 
        
    return  all_attributes



@hydra.main(config_path='conf', config_name='config', version_base='1.1')
def main(cfg: DictConfig):
    cwd = utils.get_original_cwd()
    cfg.cwd = cwd  

    # 定义要处理的根目录和结果存储目录
    root_dir = os.path.join(cwd,cfg.root_dir)  
    result_dir = os.path.join(cwd,cfg.excel_path)  
    os.makedirs(result_dir, exist_ok=True)

    # 获取根目录本身的属性
    all_attributes = [get_attributes(root_dir,root_dir)]
    # 遍历根目录下的所有文件夹和文件
    for dirpath, dirnames, filenames in os.walk(root_dir):
        dirnames.sort()
        filenames.sort() 
        # 获取当前文件夹下所有子文件夹的属性
        for dirname in dirnames:
            all_attributes.append(get_attributes(os.path.join(dirpath, dirname),root_dir))
        
        # 获取当前文件夹下所有文件的属性
        for filename in filenames:
            # all_attributes.append(get_attributes(os.path.join(dirpath, filename),root_dir))
            if filename.endswith('.docx') and os.path.getsize(os.path.join(dirpath,filename)) > 0:
               all_attributes.extend(dir_content(os.path.join(dirpath, filename),root_dir))
            else:
               all_attributes.append(get_attributes(os.path.join(dirpath, filename),root_dir))

    # 将所有属性保存到一个CSV文件
    df = pd.DataFrame(all_attributes)
    
    csv_filename = os.path.join(result_dir, 'ner_all.xlsx')
    df.index.name = 'ID'
    df.to_excel(csv_filename, index=True)

    print("所有属性已成功提取并保存到excel文件中。")

if __name__ == '__main__':
    main()
