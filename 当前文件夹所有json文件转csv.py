import json
import csv
#
# # JSON文件路径
# json_file_path = r"C:\Users\yuchenqiang\Desktop\新建文件夹 (6)\ProveType.json"
#
# # CSV文件路径
# csv_file_path = r"C:\Users\yuchenqiang\Desktop\新建文件夹 (6)\ProveType.csv"
#
# # 读取JSON文件
# with open(json_file_path, 'r', encoding='utf-8') as file:
#     data = json.load(file)
#
# # 写入CSV文件
# with open(csv_file_path, 'w', newline='', encoding='utf-8') as csvfile:
#     fieldnames = ['ProveTypeId', 'IllnessId', 'ProveType', 'Order']
#     writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
#
#     writer.writeheader()
#     for item in data:
#         writer.writerow(item)
import os
import json
import csv

# 设置包含JSON文件的文件夹路径
json_folder_path = r"C:\Users\yuchenqiang\Desktop\新建文件夹 (6)"
# 设置输出CSV文件的文件夹路径
csv_folder_path = r"C:\Users\yuchenqiang\Desktop\新建文件夹 (6)\csv_files"

# 确保输出文件夹存在
if not os.path.exists(csv_folder_path):
    os.makedirs(csv_folder_path)

# 遍历文件夹中的所有JSON文件
for filename in os.listdir(json_folder_path):
    if filename.endswith('.json'):
        # 构建完整的文件路径
        json_file_path = os.path.join(json_folder_path, filename)
        # 读取JSON文件
        with open(json_file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)

        # 构建CSV文件名和路径
        csv_filename = filename[:-5] + '.csv'  # 替换.json扩展名为.csv
        csv_file_path = os.path.join(csv_folder_path, csv_filename)

        # 写入CSV文件
        with open(csv_file_path, 'w', newline='', encoding='utf-8') as csvfile:
            # 假设所有JSON对象都有相同的键，使用第一个对象的键作为列名
            if data:
                fieldnames = data[0].keys()
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

                writer.writeheader()
                for item in data:
                    writer.writerow(item)

print("所有JSON文件已转换为CSV文件。")
