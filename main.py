import os
import sys
import xlwt
import random

# 打开界面与说明
print("--------------------------------------------------------")
print("------------欢迎使用舒尔特方格xlsx创建器 V1.5----------------")
print("--注意：本程序默认将文件创建在当前目录下,建议放在一个空文件夹中操作---")
print("------------因为遇到重名文件时，原始文件将被修改----------------")
print("---------------本次创建文件数量上限为100个-------------------")
print("----------格数上限为25*25，且必须为长宽个数相同的正方形----------")
print("--------------------------------------------------------")

# 输入需要创建的文件个数
# 需要写输入不为int的报错返回while开头
check_build_number = 0  # 创建成功为1
while check_build_number == 0:  # 输入创建个数
    try:
        number = int(input("输入需要创建的文件个数(1~100)，并点击回车(Enter):"))
    except ValueError:
        print("错误：只能输入数字，且范围在1~100")
    else:
        if 0 < number < 101:
            check_build_number = 1
        else:
            print("错误：数字超过范围(1~100),创建失败，请重新创建")

# 创建文件大小
check_build_size = 0
while check_build_size == 0:  # 输入创建个数
    try:
        size = int(input("输入需要创建长度(1~25)，并点击回车(Enter):"))
    except ValueError:
        print("错误：只能输入数字，且范围在1~25")
    else:
        if 0 < size < 26:
            check_build_size = 1
        else:
            print("错误：数字超过范围(1~25),创建失败，请重新创建")

# 创建文件
i = 1
total_size = size * size


my_style = xlwt.XFStyle()

al = xlwt.Alignment()
al.horz = 0x02 # 设置水平居中
al.vert = 0x01 # 设置垂直居中
my_style.alignment = al

# 边框格式
# 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
# 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
borders = xlwt.Borders()
borders.left = 2
borders.right = 2
borders.top = 2
borders.bottom = 2
borders.left_colour = 0x08
borders.right_colour = 8
borders.top_colour = 8
borders.bottom_colour = 8
my_style.borders = borders  # 设置边框


while i < number + 1:  # 创建number个xls文件

    arr = list(range(1, total_size + 1))
    random_number_size = total_size
    xls = xlwt.Workbook(encoding='utf-8')  # 新建工作簿
    sht1 = xls.add_sheet("Sheet1", cell_overwrite_ok=True)  # 建立Sheet

    # 随机数
    x_dimension = y_dimension = 0  # 位置
    while random_number_size != 0:
        random_number = random.randint(0, random_number_size-1)  # 生成随机数0~total_size


        space_size = sht1.col(x_dimension)
        space_size.width = 256 * 6
        sht1.row(y_dimension).height_mismatch = True
        sht1.row(y_dimension).height = 3 * 256  # 20为基准数，40意为40磅

        # 在文件中写数

        sht1.write(x_dimension, y_dimension, str(arr[random_number]),my_style)
        if random_number != random_number_size:
            arr[random_number:] = arr[random_number+1:]
        else:
            arr.pop()


        # 序列字符串中删除选中的数字
        y_dimension = y_dimension + 1
        if y_dimension == size:
            y_dimension = 0
            x_dimension = x_dimension + 1

        random_number_size = random_number_size - 1

    # 创建文件名，并保存
    name_i = '%d.xls' % i  # char联合
    xls.save(name_i)  # 保存文件
    print(name_i,"保存成功")
    # 文件数 迭代 计数
    i = i + 1

print("创建成功,请关闭程序")
os.system("pause")
