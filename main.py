import os
import sys
import xlwt
import random
# 重新整理代码，用def表示

#pyinstaller -F -i favicon.ico 舒尔特方格.py
#http://www.ico51.cn/
#ico 16*16

def Main_window_Schulte_Grid():
    # 打开界面与说明
    print("---------------------------------------------------------")
    print("------------欢迎使用舒尔特方格xlsx创建器 V2.1----------------")
    print("--注意：本程序默认将文件创建在当前目录下,建议放在一个空文件夹中操作---")
    print("------------因为遇到重名文件时，原始文件将被修改----------------")
    print("---------------本次创建表格数量上限为100个-------------------")
    print("----------格数上限为25*25，且必须为长宽个数相同的正方形----------")
    print("--------------------------------------------------------")

    # 输入需要创建的文件个数
    # 需要写输入不为int的报错返回while开头
    check_is_right = 0  # 创建成功为1
    while check_is_right == 0:  # 输入创建个数
        try:
            number = int(input("输入需要创建舒尔特方格的个数(1~100)，并点击回车(Enter):"))
        except ValueError:
            print("错误：只能输入数字，且范围在1~100")
        else:
            if 0 < number < 101:
                check_is_right = 1
            else:
                print("错误：数字超过范围(1~100),创建失败，请重新创建")

    # 创建文件大小
    check_is_right = 0
    while check_is_right == 0:  # 输入创建个数
        try:
            size = int(input("输入需要创建长度(1~25)，并点击回车(Enter):"))
        except ValueError:
            print("错误：只能输入数字，且范围在1~25")
        else:
            if 0 < size < 26:
                check_is_right = 1
            else:
                print("错误：数字超过范围(1~25),创建失败，请重新创建")

    return number,size

def set_MyExcelStyle():
    # 创建风格
    style = xlwt.XFStyle()   # 创建excel的风格中
    style.alignment.horz = 0x02  # 设置水平居中
    style.alignment.vert = 0x01  # 设置垂直居中

    # 边框格式
    # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    style.borders.left = 2
    style.borders.right = 2
    style.borders.top = 2
    style.borders.bottom = 2
    style.borders.left_colour = 0x08
    style.borders.right_colour = 8
    style.borders.top_colour = 8
    style.borders.bottom_colour = 8
    return style

def Creat_Excel_File(Number, Size):
    xls = xlwt.Workbook(encoding='utf-8')  # 新建工作簿
    for i in range(1, Number+1):  # 创建Number个页面,range取前不取后
        arr = list(range(1, (Size * Size) + 1))  # 建立填入数字候选数组,
        sht1 = xls.add_sheet("Schulte_Grid_No." + str(i), cell_overwrite_ok=True)  # 建立Sheet

        # 随机数
        for x_dimension in range(0, Size):
            for y_dimension in range(0, Size):
                # 表格长宽
                sht1.col(x_dimension).width_mismatch = True
                sht1.col(x_dimension).width = 256 * 6
                sht1.row(y_dimension).height_mismatch = True
                sht1.row(y_dimension).height = 256 * 3  # 20为基准数，40意为40磅

                # 在文件中写数
                random_number = random.randint(0, len(arr) - 1)  # 生成随机数
                sht1.write(x_dimension, y_dimension, str(arr[random_number]), set_MyExcelStyle())

                # 序列字符串中删除选中的数字
                del arr[random_number]

    # 创建文件名，并保存
    xls.save('舒尔特方格.xls')  # 保存文件
    print('舒尔特方格.xls', "保存成功,请关闭程序")



if __name__ == '__main__':
    number, size = Main_window_Schulte_Grid()
    Creat_Excel_File(number, size)
    os.system("pause")
