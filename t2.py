# -*- coding: utf-8 -*-
import arcpy
import os
import csv
import codecs
import sys
reload(sys)  # 重新加载sys模块以修改默认编码
sys.setdefaultencoding('utf-8')  # 设置为UTF-8编码

# 设置目录
shp_dir = ur"D:/GISData/FZ/市区"

# 输出 CSV
output_csv = "arcpy_field_output.csv"
csvfile = codecs.open(output_csv, "w", encoding="gbk")
writer = csv.writer(csvfile)

# 写表头（英文）
writer.writerow(["Filename", "FieldName", "FieldAlias", "Type", "Length", "Scale"])

# 定义数值类型集合
numeric_types = {"Double", "Single", "Integer", "SmallInteger", "BigInteger"}


# 遍历所有 SHP 文件
for root, dirs, files in os.walk(shp_dir):
    for file in files:
        if file.lower().endswith(".shp"):
            shp_path = os.path.join(root, file)
            try:
                fields = arcpy.ListFields(shp_path)
                for idx, field in enumerate(fields):
                    # 字段长度或精度判断
                    if field.type in numeric_types:
                        row = [
                            file,
                            field.name,
                            field.aliasName,
                            field.type,
                            field.precision,
                            field.scale
                        ]
                    else:
                        row = [
                            file,
                            field.name,
                            field.aliasName,
                            field.type,
                            field.length,
                            field.scale
                        ]

                    writer.writerow([unicode(s) for s in row])
            except Exception as e:
                writer.writerow([file, "[Error]", str(e)])

csvfile.close()
print(output_csv)
