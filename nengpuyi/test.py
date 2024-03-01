#
# import win32com.client
#
# def read_floating_text_boxes(doc_file):
#     word = win32com.client.Dispatch("Word.Application")
#     doc = word.Documents.Open(doc_file)
#     for shape in doc.Shapes:
#         if shape.Type == 17:  # 17 表示文本框类型
#             text = shape.TextFrame.TextRange.Text
#             print(text.encode('GBK'))
#     doc.Close()
#     word.Quit()
#
#
#
#
# if __name__ == '__main__':
#     # 调用函数并传入 Word 文档的路径
#     read_floating_text_boxes("E:/金现代/凯盛/设备调研表/物化所设备数据/仪器原始数据/能谱仪SEM-EDS/42-0101.doc")


# import win32com.client
#
#
# def read_text_boxes(doc_file):
#     # 创建Word应用程序对象
#     word = win32com.client.Dispatch("Word.Application")
#     # 打开文档
#     doc = word.Documents.Open(doc_file)
#
#     # 遍历文档中的所有文本框
#     for shape in doc.Shapes:
#         if shape.Type == 17:  # Type为17表示文本框
#             # 获取文本框的文本内容
#             text = shape.TextFrame.TextRange.Text
#             # 解析文本框中的数据
#             parsed_data = parse_text(text)
#             # 输出解析后的数据
#             for key, values in parsed_data.items():
#                 print(key + ":", values)
#
#     # 关闭文档和Word应用程序
#     doc.Close()
#     word.Quit()
#
#
# def parse_text(text):
#     # 将文本按行分割成列表
#     lines = text.strip().split('\n')
#
#     # 创建一个空字典来存储数据
#     data = {}
#
#     # 解析每一行的数据
#     for line in lines:
#         # 如果当前行不为空
#         if line.strip():
#             # 根据制表符分割每一行，并去除首尾空格
#             parts = [part.strip() for part in line.split('\t')]
#             # 提取键值对
#             key = parts[0]
#             values = parts[1:]
#             # 将键值对添加到字典中
#             data[key] = values
#
#     return data
#
#
# # 调用函数并传入.doc文件的路径
# doc_file = "E:/金现代/凯盛/设备调研表/物化所设备数据/仪器原始数据/能谱仪SEM-EDS/42-0101.doc"
# read_text_boxes(doc_file)


# import win32com.client
#
# def read_text_boxes(doc_file):
#     # 创建Word应用程序对象
#     word = win32com.client.Dispatch("Word.Application")
#     # 打开文档
#     doc = word.Documents.Open(doc_file)
#
#     # 遍历文档中的所有文本框
#     for shape in doc.Shapes:
#         if shape.Type == 17:  # Type为17表示文本框
#             # 获取文本框的文本内容
#             text = shape.TextFrame.TextRange.Text
#             print("文本框内容:", repr(text))  # 打印文本框内容，以便调试
#
#     # 关闭文档和Word应用程序
#     doc.Close()
#     word.Quit()
#
# # 调用函数并传入.doc文件的路径
# doc_file = "E:/金现代/凯盛/设备调研表/物化所设备数据/仪器原始数据/能谱仪SEM-EDS/42-0101.doc"
# read_text_boxes(doc_file)


# import re
#
#
# def parse_text_box_content(text):
#     # 使用正则表达式匹配重量、原子、化合物、化学式以及百分比的数据
#     pattern = r'([A-Za-z\s]+)\r\x07([\d.]+)\r\x07([\d.]+)\r\x07([\d.]+)\r\x07([A-Za-z0-9\s]+)\r\x07'
#     matches = re.findall(pattern, text)
#
#     data = []
#     # 将匹配到的数据转换成字典形式存储
#     for match in matches:
#         element_data = {
#             "元素": match[0].strip(),
#             "重量": float(match[1]),
#             "原子": float(match[2]),
#             "化合物": float(match[3]),
#             "化学式": match[4].strip()
#         }
#         data.append(element_data)
#
#     return data
#
#
# # 测试解析函数
# text = '元素\r\x07重量\r\x07原子\r\x07化合物\r\x07化学式\r\x07\r\x07\r\x07   \r\x07百分比\r\x07百分比\r\x07百分比\r\x07   \r\x07\r\x07\r\x07C K\r\x0727.25\r\x0733.30\r\x0799.84\r\x07CO2\r\x07\r\x07\r\x07Ca K\r\x070.12\r\x070.04\r\x070.16\r\x07CaO\r\x07\r\x07\r\x07O\r\x0772.64\r\x0766.65\r\x07\r\x07\r\x07\r\x07\r\x07总量\r\x07100.00\r\x07\r\x07\r\x07\r\x07\r\x07\r\x07\r\r'
# parsed_data = parse_text_box_content(text)
# print(parsed_data)


# import win32com.client
# import re
#
#
# def read_text_boxes(doc_file):
#     # 创建Word应用程序对象
#     word = win32com.client.Dispatch("Word.Application")
#     # 打开文档
#     doc = word.Documents.Open(doc_file)
#     texts = []
#     # 遍历文档中的所有文本框
#     for shape in doc.Shapes:
#         if shape.Type == 17:  # Type为17表示文本框
#             # 获取文本框的文本内容
#             text = shape.TextFrame.TextRange.Text.strip()
#             print("文本框内容:", repr(text))  # 打印文本框内容，以便调试
#             texts.append(text)
#     # 关闭文档和Word应用程序
#     doc.Close()
#     word.Quit()
#     return texts
#
#
# def parse_text_box_content(texts):
#     # 使用正则表达式匹配重量、原子、化合物、化学式以及百分比的数据
#     pattern = r'([A-Za-z\s]+)\r\x07([\d.]+)\r\x07([\d.]+)\r\x07([\d.]+)\r\x07([A-Za-z0-9\s]+)\r\x07'
#     data = []
#     # 遍历文本框内容
#     for text in texts:
#         matches = re.findall(pattern, text)
#         # 将匹配到的数据转换成字典形式存储
#         for match in matches:
#             element_data = {
#                 "元素": match[0].strip(),
#                 "重量": float(match[1]),
#                 "原子": float(match[2]),
#                 "化合物": float(match[3]),
#                 "化学式": match[4].strip()
#             }
#             data.append(element_data)
#     return data
#
# if __name__ == '__main__':
#     # 调用函数并传入 Word 文档的路径
#     doc_file = "E:/金现代/凯盛/设备调研表/物化所设备数据/仪器原始数据/能谱仪SEM-EDS/42-0101.doc"
#     texts = read_text_boxes(doc_file)
#     parsed_data = parse_text_box_content(texts)
#     print(parsed_data)



# import win32com.client
# def read_text_boxes(doc_file):
#     # 创建Word应用程序对象
#     word = win32com.client.Dispatch("Word.Application")
#     # 打开文档
#     doc = word.Documents.Open(doc_file)
#     text_boxes = []
#     # 遍历文档中的所有文本框
#     for shape in doc.Shapes:
#         if shape.Type == 17:  # Type为17表示文本框
#             # 获取文本框的文本内容并进行处理
#             text = shape.TextFrame.TextRange.Text.strip()
#             # 替换掉文本中的'\r'换行符和'\x07'响铃字符
#             text = text.replace('\r', '\n').replace('\x07', '')
#             # 将文本框内容按行分割，并移除空行和只包含空白字符的行
#             lines = [line.strip() for line in text.splitlines() if line.strip()]
#             # 将字典添加到列表中
#             text_boxes.append(lines)
#     # 关闭文档和Word应用程序
#     doc.Close()
#     word.Quit()
#     return text_boxes
#
# if __name__ == '__main__':
#     # 调用函数并传入 Word 文档的路径
#     doc_file = "E:/金现代/凯盛/设备调研表/物化所设备数据/仪器原始数据/能谱仪SEM-EDS/42-0101.doc"
#     text_boxes_content = read_text_boxes(doc_file)
#     for content in text_boxes_content:
#         print(content)

# import win32com.client
#
# def read_text_boxes(doc_file):
#     # 创建Word应用程序对象
#     word = win32com.client.Dispatch("Word.Application")
#     # 打开文档
#     doc = word.Documents.Open(doc_file)
#     text_boxes_content = []
#     # 遍历文档中的所有文本框
#     for shape in doc.Shapes:
#         if shape.Type == 17:  # Type为17表示文本框
#             # 获取文本框的文本内容并进行处理
#             text = shape.TextFrame.TextRange.Text.strip()
#             # 替换掉文本中的'\r'换行符和'\x07'响铃字符
#             # 替换特殊字符
#             text = text.replace('\r\x07\r\x07\r\x07', '\n').replace('\r\x07', '\t').replace('\r', '\n').replace('\x07', '')
#
#             # 将文本框内容按行分割，并移除空行和只包含空白字符的行
#             lines = [line.strip() for line in text.splitlines() if line.strip()]
#             # 将字典添加到列表中
#             text_boxes_content.append(lines)
#     # 关闭文档和Word应用程序
#     doc.Close()
#     word.Quit()
#     return text_boxes_content
#
# def print_formatted_content(text_boxes_content):
#     max_lengths = [max(len(row[i]) for row in text_boxes_content) for i in range(len(text_boxes_content[0]))]
#     for content in text_boxes_content:
#         for i, line in enumerate(content):
#             items = line.split('\t')
#             formatted_line = ''.join([items[j].ljust(max_lengths[j] + 4) for j in range(len(items))])
#             print(formatted_line)
#             if i == 0:  # Add an empty line after the first row
#                 print()
#         print()
#
# if __name__ == '__main__':
#     # 调用函数并传入 Word 文档的路径
#     doc_file = "E:/金现代/凯盛/设备调研表/物化所设备数据/仪器原始数据/能谱仪SEM-EDS/42-0101.doc"
#     text_boxes_content = read_text_boxes(doc_file)
#     print_formatted_content(text_boxes_content)




import win32com.client
import re

def read_text_boxes(doc_file):
    # 创建Word应用程序对象
    word = win32com.client.Dispatch("Word.Application")
    # 打开文档
    doc = word.Documents.Open(doc_file)
    text_boxes_content = []
    # 遍历文档中的所有文本框
    for shape in doc.Shapes:
        if shape.Type == 17:  # Type为17表示文本框
            # 获取文本框的文本内容并进行处理
            text = shape.TextFrame.TextRange.Text.strip()
            # 替换掉文本中的'\r'换行符和'\x07'响铃字符
            # 替换特殊字符
            text = text.replace('\r\x07\r\x07\r\x07', '\n').replace('\r\x07', '\t').replace('\r', '\n').replace('\x07', '')

            # 将文本框内容按行分割，并移除空行和只包含空白字符的行
            lines = [line.strip() for line in text.splitlines() if line.strip()]

            # 添加处理字符串的步骤：根据冒号分割字符串，并存放在元组中
            split_lines = []
            for line in lines:
                split_line = re.split(r'\s*:\s*', line, maxsplit=1)
                if len(split_line) == 2:
                    split_lines.append(tuple(split_line))
                else:
                    # 处理列表部分，使用换行符进行分割
                    split_lines.extend([(item.strip(), '') for item in line.split('\n') if item.strip()])
            # 将元组添加到列表中
            text_boxes_content.append(split_lines)
    # 关闭文档和Word应用程序
    doc.Close()
    word.Quit()
    return text_boxes_content

def print_formatted_content(text_boxes_content):
    max_lengths = [max(len(row[i][0]) for row in text_boxes_content) for i in range(len(text_boxes_content[0]))]
    for content in text_boxes_content:
        for i, line in enumerate(content):
            if len(line) == 2:
                formatted_line = ''.join([line[0].ljust(max_lengths[0] + 4), line[1]])
            else:
                formatted_line = line[0]
            print(formatted_line)
            if i == 0:  # Add an empty line after the first row
                print()
        print()

if __name__ == '__main__':
    # 调用函数并传入 Word 文档的路径
    doc_file = "E:/金现代/凯盛/设备调研表/物化所设备数据/仪器原始数据/能谱仪SEM-EDS/42-0101.doc"
    text_boxes_content = read_text_boxes(doc_file)
    print_formatted_content(text_boxes_content)















