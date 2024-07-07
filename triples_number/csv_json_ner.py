import csv
import json

# 输入CSV文件名
csv_re = 'triples_number/relation.csv'
csv_ner = 'triples_number/ner.csv'
# 输出JSON文件名
json_re = 'triples_number/relation.json'
json_ner = 'triples_number/ner.json'

# 初始化一个空的字典，用于存储最终的JSON数据
json_links = {"links": []}
json_nodes = {"nodes": []}
json_item = {}

# 读取关系CSV文件
with open(csv_re, mode='r', newline='', encoding='utf-8') as csv_file:
    csv_reader = csv.reader(csv_file)
    for row in csv_reader:
        source, target, relation = row
        json_links["links"].append({
            "relation": relation,
            "source": int(source),
            "target": int(target)
        })

# 读取NER CSV文件
with open(csv_ner, mode='r', newline='', encoding='utf-8') as csv_file:
    csv_reader = csv.reader(csv_file)
    
    for row in csv_reader:
        number, name, parent_path, path, time, type, depth, title, content, parent_title = row
        
        json_item[number] = {
            "name": name,
            "parent_path": parent_path,
            "path": path,
            "time": time,
            "type": type,
            "content": content.strip(),
        }

        json_nodes["nodes"].append({
            "index": int(number),
            "name": name,
            "type": type,
            "depth": int(depth),
        })

# 合并链接和节点数据
json_combined = {**json_links, **json_nodes}

# 将链接和节点数据写入JSON文件
with open(json_re, mode='w', encoding='utf-8') as json_file:
    json.dump(json_combined, json_file, ensure_ascii=False, indent=4)

# 将NER数据写入JSON文件
with open(json_ner, mode='w', encoding='utf-8') as json_file:
    json.dump(json_item, json_file, ensure_ascii=False, indent=4)

print("JSON文件已生成！")

