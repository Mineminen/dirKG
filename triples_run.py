import os
import pandas as pd
import time
from omegaconf import DictConfig
from docx import Document
import hydra
from hydra import utils
import re



# 从ner_all.csv中读取路径和编号映射关系
def load_number_mapping(ner_all_path):

    df = pd.read_excel(ner_all_path,dtype={'Path':str,'Title':str})
   
    df['path_title'] = df.apply(lambda row: os.path.join(str(row['Path']), str(row['Title'])), axis=1)
    number_mapping_title = dict(zip(df['path_title'], df['ID']))
    number_mapping=dict(zip(df['Path'], df['ID']))
    return number_mapping_title,number_mapping


  
def get_attributes(path, number_mapping,root_dir):
    # 获取从根目录到当前路径的相对路径
    relpath_path = os.path.relpath(path, os.path.dirname(root_dir))
    # 获取当前路径的编号
    number = number_mapping.get(path, None)

    try:
        stat = os.stat(path)
        parent_dir = '' if path == root_dir else os.path.dirname(relpath_path)
        # 计算当前路径相对于root_dir的深度
        depth = relpath_path.count(os.sep) + 1
        
        attributes = {
            'Number': number,
            'Name': os.path.basename(path),
            'Parent': parent_dir,
            'Path': relpath_path,
            'Creation Time': time.ctime(stat.st_ctime),
            'Label': 'Directory' if os.path.isdir(path) else 'File' ,
            'Depth': depth ,
            'Title': None,
            'Content': None,
            'Parent_title': None, 
        }
        
        if os.path.isfile(path):
            if path.endswith('.txt') or path.endswith('.py'):
                with open(path, 'r', encoding='utf-8') as file:
                    attributes['Content'] = file.read()
        return attributes
    except Exception as e:
        print(f"Error processing {path}: {e}")
        return {'Number': number, 'Path': os.path.basename(path), 'Creation Time': None,
                'Label': None, 'Depth': None, 'Content': f"Error processing file: {e}"}
    
def dir_content(doc_path, root_dir,number_mapping_title,number_mapping):
    all_attributes = []
    relationships = []
    
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
    path_title=os.path.join(doc_path,doc_name)
    number = number_mapping_title.get(path_title, None)
    # 栈，存储标题和标题级别
    stack = [(number, doc_name, 0, [])]
    
    for paragraph in document.paragraphs:
        level = get_heading_level(paragraph)
        if level is not None:
            current_heading = paragraph.text.strip()
            stack[-1][3].extend(current_content)

            current_content = []
            doc_path_title= os.path.join(doc_path,current_heading)
            number = number_mapping_title.get(doc_path_title, None)
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
        attribute=get_attributes(doc_path,number_mapping,root_dir)
        current_content = ','.join(current_content)
        # 查找父标题
        parent_title = ''
        for j in range(i - 1, -1, -1):
            if stack[j][2] < current_level:
                parent_title = stack[j][1]
                break
        if parent_title:
            doc_path_fa_title=os.path.join(doc_path,parent_title) 
            parent_number=number_mapping_title.get(doc_path_fa_title, None)
            relationships.append((parent_number,current_number,'include'))
        else:
            attr=get_attributes(os.path.dirname(doc_path),number_mapping,root_dir)
            parent_number=attr['Number']
            relationships.append((parent_number,current_number,'include'))
        attribute['Number']=current_number
        attribute['Content']=current_content
        attribute['Title']=current_title
        attribute['Parent_title']=parent_title

        all_attributes.append(attribute) 
                  
    return  all_attributes,relationships
   
    


@hydra.main(config_path='conf', config_name='config', version_base='1.1')
def main(cfg: DictConfig):
    cwd = utils.get_original_cwd()
    cfg.cwd = cwd
    # 定义要处理的根目录和结果存储目录
    root_dir = os.path.join(cwd, cfg.root_dir)
    result_dir = os.path.join(cwd, cfg.triple_number_path)
    os.makedirs(result_dir, exist_ok=True)
    # 读取ner_all.xlsx中的路径和编号映射关系
    ner_all_path = os.path.join(cwd, 'excel','ner_all.xlsx')
    number_mapping_title,number_mapping = load_number_mapping(ner_all_path)
    # 定义结果文件路径
    csv_filename_ner = os.path.join(result_dir, f'ner.csv')
    csv_filename_relation = os.path.join(result_dir, f'relation.csv')
    # 在循环开始之前删除文件
    if os.path.exists(csv_filename_ner):
        os.remove(csv_filename_ner)

    if os.path.exists(csv_filename_relation):
        os.remove(csv_filename_relation)
    # 开始处理根目录
    root_attributes = [get_attributes(root_dir, number_mapping,root_dir)]
    df_ner = pd.DataFrame(root_attributes)
    df_ner.to_csv(csv_filename_ner, mode='a',index=False, header=False)
    # 获取当前目录下所有子文件夹和文件的属性
    for dirpath, dirnames, filenames in os.walk(root_dir):
        if not dirnames and not filenames:
           continue
        dirnames.sort()
        filenames.sort()
        all_attributes_tail = []
        relationships = []
        # 获取子文件夹的属性
        all_attributes_head = [get_attributes(dirpath, number_mapping,root_dir)]
        head = all_attributes_head[0]['Number']
        for dirname in dirnames:   
            sub_dirpath = os.path.join(dirpath, dirname)
            tail_attributes = get_attributes(sub_dirpath, number_mapping,root_dir)
            all_attributes_tail.append(tail_attributes)
            tail = tail_attributes['Number']
            relationships.append((head, tail, 'include'))

        # 获取文件的属性
        for filename in filenames:
            filepath = os.path.join(dirpath, filename)
           
            if  filepath.endswith('.docx') and os.path.getsize(filepath) > 0:
                some_attributes,some_relationships =dir_content(filepath,root_dir,number_mapping_title,number_mapping)
                all_attributes_tail.extend(some_attributes)
                relationships.extend(some_relationships)
                
            else:
                tail_attributes = get_attributes(filepath, number_mapping,root_dir)
                tail = tail_attributes['Number']
                relationships.append((head, tail, 'include'))
                all_attributes_tail.append(tail_attributes)


        df_tail = pd.DataFrame(all_attributes_tail)
        df_tail.to_csv(csv_filename_ner, mode='a',index=False, header=False)

        df_relation = pd.DataFrame(relationships, columns=['head', 'tail', 'Relation'])
        df_relation.to_csv(csv_filename_relation, mode='a',index=False, header=False)
    print("所有属性已成功提取并保存到CSV文件中。")

if __name__ == '__main__':
    main()
