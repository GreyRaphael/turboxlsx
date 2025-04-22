from turboxlsx import BookWriter
import time
from fastxlsx import DType, WriteOnlyWorkbook, WriteOnlyWorksheet, write_many


COL_COUNT = 10
ROW_COUNT = 65535

data = {
    "sheet你好1": {
        "headers": [f"col-第{i + 1}列" for i in range(COL_COUNT)],
        "data2d": [[float(j + i) for i in range(ROW_COUNT)] for j in range(COL_COUNT)],
    },
    "sheet你好2": {
        "headers": [f"col-第{i + 1}列" for i in range(COL_COUNT + 10)],
        "data2d": [[float(j + i) for i in range(ROW_COUNT)] for j in range(COL_COUNT + 10)],
    },
}

start = time.time()
book = BookWriter()
book.add_sheet("sheet1", data["sheet你好1"]["headers"])
for i, col in enumerate(data["sheet你好1"]["data2d"]):
    book.add_column_number(0, col)
book.add_sheet("sheet2", data["sheet你好2"]["headers"])
for i, col in enumerate(data["sheet你好2"]["data2d"]):
    book.add_column_number(1, col)

book.save(name="test.xlsx")
print(time.time() - start)

data = {
    "sheet你好1": {
        "headers": [f"col-第{i + 1}列" for i in range(COL_COUNT)],
        "data2d": [[i + j for j in range(COL_COUNT)] for i in range(ROW_COUNT)],
    },
    "sheet你好2": {
        "headers": [f"col-第{i + 1}列" for i in range(COL_COUNT + 10)],
        "data2d": [[i + j for j in range(COL_COUNT + 10)] for i in range(ROW_COUNT)],
    },
}
print("data prepared")

start = time.time()
wb = WriteOnlyWorkbook()
ws = wb.create_sheet("sheet1")
ws.write_matrix((0, 0), data["sheet你好1"]["data2d"])
ws2 = wb.create_sheet("sheet2")
ws2.write_matrix((0, 0), data["sheet你好2"]["data2d"])
wb.save("fast.xlsx")
print(time.time() - start)
