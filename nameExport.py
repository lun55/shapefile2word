# -*- coding: utf-8 -*-
import arcpy
import os
import csv

# 设置目标目录（包含多个 Shapefile）
shp_dir = r"D:\GISData\FZ\OSM\shape"

# 输出 CSV 文件路径（使用 UTF-8 编码）
output_csv = "field_output.csv"

# 打开 CSV 文件写入
with open(output_csv, "wb") as csvfile:
    writer = csv.writer(csvfile)
    # 英文表头，不需要 encode
    writer.writerow([
        "Filename",
        "FieldName",
        "FieldAlias",
        "Type",
        "Length",
        "Scale"
    ])
    
    for root, dirs, files in os.walk(shp_dir):
        for file in files:
            if file.lower().endswith(".shp"):
                shp_path = os.path.join(root, file)
                try:
                    fields = arcpy.ListFields(shp_path)
                    for idx, field in enumerate(fields):
                        row = [
                            file.encode("utf-8"),
                            field.name.encode("utf-8"),
                            field.aliasName.encode("utf-8"),
                            field.type.encode("utf-8"),
                            str(field.length),
                            str(field.scale)
                        ]
                        writer.writerow(row)
                except Exception as e:
                    writer.writerow([file.encode("utf-8"), "[错误]", str(e).encode("utf-8")])

print("字段信息已导出到 {}".format(output_csv))
