import jieba
import xlwt, xlrd
import xlwings as xw
from collections import Counter

# 定义一个空列表
all_word_list = []


# 分词
def trans_CN(text):
    # 接收分词的字符串
    word_list = jieba.cut(text)
    # 分词后在单独个体之间加上空格
    result = " ".join(word_list)
    # 转换成list
    result = result.split(" ")
    return result


# 判断词是否为中文
def is_Chinese(word):
    for ch in word:
        if '\u4e00' <= ch <= '\u9fff':
            return True
    return False


start_col = 2  # 处理Excel文件开始列
end_col = 52  # 处理Excel结束列

# 指定不显示地打开Excel，读取Excel文件
app = xw.App(visible=False, add_book=False)
wb = app.books.open(r"C:\Users\20610\PycharmProjects\毕业论文爬虫项目\关键词频.xlsx")  # 打开Excel文件
sheet = wb.sheets[9]  # 选择第0个表单

# 读取Excel表单前52列的数据
for col in range('A', 'Z'):
    print(col)
    col_str = 'A'
    # 循环中引用Excel的sheet和range的对象，读取C列的每一行的值
    content_text = sheet.range().value
    # print(content_text)
    if not content_text:
        continue
    if not isinstance(content_text, str):
        continue
    # 长度小于4的语句 过滤
    if len(content_text) > 3:
        word_list = trans_CN(content_text)
        print("分词后", word_list)
        # 判断列表元素是否为中文，将非中文词移除
        for s in word_list:
            if not is_Chinese(s):
                word_list.remove(s)
        all_word_list += word_list

# 统计列表中元素出现的频率
counter = Counter(all_word_list)
print("统计频率完成")

# 将列表中的元素按照频率大小排序
result_list = sorted(counter.items(), key=lambda x: x[1], reverse=True)

# 将结果写入表格
print("开始写入表格")
myWorkbook = xlwt.Workbook()
mySheet = myWorkbook.add_sheet('Sheet1', cell_overwrite_ok=True)
rows = 0
for i in result_list:
    mySheet.write(rows, 0, i[0])
    mySheet.write(rows, 1, i[1])
    rows += 1
myWorkbook.save('result.xls')

# 保存并关闭Excel文件
wb.save()
wb.close()
