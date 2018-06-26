# 查询不同的语句并导出为Excel表格应用(一条查询对应一个sheet)：
import sys
import xlwt
import MySQLdb


def write(results, fields, sheet):
    """写入字段信息和结果"""
    for field in range(0, len(fields)):
        sheet.write(0, field, fields[field][0])

    row = 1
    col = 0
    for row in range(1, len(results)+1):
        for col in range(0, len(fields)):
            sheet.write(row, col, u'%s' % results[row-1][col])


def search_data(cursor, sql):
    """获取数据结果和字段信息"""
    cursor.execute(sql)
    results = cursor.fetchall()
    fields = cursor.description
    return results, fields


db = MySQLdb.connect(host="主机", port=端口, user="用户名", password="密码", db="xxx", charset='utf8')

cursor1 = db.cursor()
# cursor2 = db.cursor()
# cursor3 = db.cursor()

sql1 = """sql查询语句1"""
# sql2 = """sql查询语句2"""
# sql3 = """sql查询语句3"""

results1, fields1 = search_data(cursor1, sql1)
# results2, fields2 = search_data(cursor2, sql2)
# results3, fields3 = search_data(cursor3, sql3)

workbook = xlwt.Workbook()
sheet1 = workbook.add_sheet('sheet_name', cell_overwrite_ok=True)
# sheet2 = workbook.add_sheet('test_data2', cell_overwrite_ok=True)
# sheet3 = workbook.add_sheet('test_data3', cell_overwrite_ok=True)

write(results1, fields1, sheet1)
# write(results2, fields2, sheet2)
# write(results3, fields3, sheet3)

workbook.save(r'./file_name.xls')

db.close()





















