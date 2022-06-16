import xlrd
import xlwt
import os
import ast


def duplicate_removal():  # 去重
    print('开始去重...')
    path = "sn.txt"
    new_list = []
    for line in open('sn.txt', 'r+'):
        new_list.append(line)
    new_list2 = list(set(new_list))  # 去重
    new_list2.sort(key=new_list.index)  # 以原list的索引为关键词进行排序
    new_txt = ''.join(new_list2)  # 将新list连接成一个字符串
    with open('newsn.txt', 'w') as f:
        f.write(new_txt)
        f.close()
    print('资产去重完毕')
    if os.path.exists(path):
        os.remove(path)

def find_data():
    data = xlrd.open_workbook('202204.xls') #资产总表，财务给
    data2 = xlrd.open_workbook("sn.xls") #待查找的序列号表
    finddata = xlwt.Workbook(encoding='utf-8')
    sheets = data.sheet_by_index(0)
    sheets2 = data2.sheet_by_index(0)
    sheets3 = finddata.add_sheet("表1",cell_overwrite_ok=True)
    col = ("资产编号","序列号")
    for i in range(0,2):
        sheets3.write(0,i,col[i])
    finddata.save("test.xls")
    alldata = []
    xlsx1_row = sheets.nrows
    xlsx2_row = sheets2.nrows
    for i in range(4,xlsx1_row):
        key_zichanbianhao = sheets.cell(i,0).value
        key_zichanquleihao = sheets.cell(i,2).value
        data1 = {"资产编号":key_zichanbianhao,"设备序列号":key_zichanquleihao}
        alldata.append(data1)
    for k in range(0,xlsx2_row):
        count = 1
        key_sn = sheets2.cell(k,0).value
        for k2 in alldata:
            if str(key_sn) in str(k2):
                file1 = open ("bianhao.txt","a",encoding="utf-8")
                file3 = open ("jilu.txt.txt","a",encoding="utf-8")
                file1.write(str(k2.get("资产编号"))+"\n")
                file3.write(str(k2)+"\n")
                file1.close()
                file3.close()
            else:
                count = count + 1 #记录for循环次数
                continue
        if count == int(xlsx1_row) - 3: #for循环次数满，即为没有查找到数据
            file2 = open("notfound.txt","a",encoding="utf-8")
            file2.write(str(key_sn)+"\n")
            file2.close()
            print(str(key_sn)+"没找到")
    foundc = 0
    foundc2 = 0
    notfoundc = 0
    with open("jilu.txt.txt","r",encoding="utf-8") as data3:
        for line in data3:
            ld = ast.literal_eval(line) #str转字典
            foundc = foundc + 1
            sheets3.write(foundc,0,ld.get("资产编号"))
            finddata.save("test.xls")
    with open("jilu.txt.txt", "r", encoding="utf-8") as data8:
        for line2 in data8:
            ld = ast.literal_eval(line2)
            foundc2 = foundc2 + 1
            sheets3.write(foundc2, 1, ld.get("设备序列号"))
            finddata.save("test.xls")
    try:
        with open("notfound.txt","r",encoding="utf-8") as data4:
            for line in data4:
                notfoundc = notfoundc + 1
    except Exception as err:
        pass

    print("查询资产总数:"+str(xlsx2_row)+","+"查到资产编号数量:"+str(foundc)+","+"未查到资产编号数量:"+str(notfoundc))

find_data()


