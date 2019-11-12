#coding=gbk
import pymysql
from openpyxl import load_workbook
from openpyxl import Workbook
import os
import uuid
from time import strftime, localtime

#获取个性化最大列数，只读取第x行，最大列数为第一个空白cell
def personal_max_col(sheet,beginRow):
    count = 1
    while sheet.cell(row=beginRow, column=count).value != None:
        count = count + 1
    return count-1

#获取个性化最大列数，只读取第一行，最大行数为第一个字段为 string
def personal_max_row(sheet,string):
    maxRowToCheck = 5000
    count = 1
    value = sheet.cell(row=count, column=1).value
    while count < maxRowToCheck and string not in str(value):
        count = count + 1
        value = sheet.cell(row=count, column=1).value
        if value == None:
            value = ""
    return count

def file_name(file_dir):
    filesList = []
    for root, dirs, files in os.walk(file_dir):
        # print(root)  # 当前目录路径
        # print(dirs)  # 当前路径下所有子目录
        # print(files)  # 当前路径下所有非目录子文件
        filesList = files
    return filesList

def getProductValue(name):
    if "国际条码" in name:
        return "product_code"#"国际条码"
    elif "商品名称"in name:
        return "product"#"商品名称"
    elif "单位" in name:
        return "unit"  # "单位"
    elif "规格" in name:
        return "specifications"  # "规格"
    elif "调拨门店"in name:
        return "store_name"  # "调拨门店"
    elif "实际可调拨数量"in name:
        return "real_trans_num"  # "实际可调拨数量"
    elif "调拨数量" in name:
        return "trans_num"  # "调拨数量"
    return -1

def getOtherValue(name):
    if "出库类型" in name:
        return "output_type"#"出库类型"
    elif "出库地点"in name:
        return "output_place"#"出库地点"
    elif "出库联系人电话" in name:
        return "output_contact"  # "出库联系人电话"
    elif "出库联系人" in name:
        return "output_person"  # "出库联系人"
    elif "到货地点"in name:
        return "input_place"#"到货地点"
    elif "联系人姓名"in name:
        return "input_person"#"联系人姓名"
    elif "联系人电话"in name:
        return "input_contact"#"联系人电话"
    elif "要求到达时间"in name:
        return "arrive_time"#"要求到达时间"
    return -1

def getWMInfo(name):
    if "合计" in name:
        return "quantity"
    elif "物流" in name:
        return "logistic"
    elif "车型" in name:
        return "car_type"
    elif name is None:
        return "area_size"
    else:
        return -1

def connectDB():
    # 打开数据库连接
    db = pymysql.connect("localhost", "root", "root", "rexuan", charset='utf8')
    return db

def closeDB(db):
    # 关闭数据库连接
    db.close()

root = "C:/Users/admin/Desktop/toInsert/"
root = "C:/Users/admin/Desktop/toInsert/test/"
filesNameList = file_name(root)


db = connectDB()
# 使用cursor()方法获取操作游标
cursor = db.cursor()

for order_name in filesNameList:
    readFilepath = root + order_name
    created_time = strftime("%Y%m%d%H%M%S", localtime())

    # 默认可读写，若有需要可以指定write_only和read_only为True
    wb = load_workbook(readFilepath)

    sheetNamesList = wb.sheetnames

    orders_id = ''.join(str(uuid.uuid4()).split('-'))
    orderSql = "INSERT INTO orders (orders_id,order_name,created_time,	quantity,logistic,area_size,car_type)VALUES	('"+orders_id+"','"+order_name+"','"+created_time+"',"


    if "仓库填写" in sheetNamesList:
        # print("开始获取仓库填写信息")
        actualSheet = wb["仓库填写"]
        maxRow = personal_max_row(actualSheet,"仓库填写完毕后邮件回发至相关人员")-1
        beginRow = 1
        # print(maxRow)
        # print(beginRow)

        tmpIndex = 1
        while tmpIndex <= maxRow:
            orderSql = orderSql + "'" + str(actualSheet.cell(row=tmpIndex,column=2).value) + "',"
            tmpIndex = tmpIndex + 1

        orderSql = orderSql[:-1]+")"
        sheetNamesList.remove("仓库填写")

    print(orderSql)
    # print(sheetNamesList)

    # try:
    #         # print(orderSql)
    #         affectRows = cursor.execute(orderSql)
    #         # print(affectRows)
    #         db.commit()
    # except:
    #         print("==order sql error:=="+ readFilepath + "====")

    #获取每个sheet的内容
    for sheetName in sheetNamesList:
        stores_sheet_id = ''.join(str(uuid.uuid4()).split('-'))

        detailSql = "INSERT INTO detail_info (id,stores_sheet_id,orders_id,"
        storesSheetSql = "INSERT INTO stores_sheet(stores_sheet_id,orders_id,quantity,"
        actualSheet = wb[sheetName]

        beginRow = personal_max_row(actualSheet,"国际条码")
        # print(beginRow)

        maxCol = personal_max_col(actualSheet,beginRow=beginRow)

        tmpCol = 1
        while tmpCol <= maxCol:
            detailSql = detailSql + str(getProductValue(actualSheet.cell(row=beginRow, column=tmpCol).value))+","
            tmpCol = tmpCol + 1

        detailSql = detailSql[:-1]
        detailSql = detailSql + ") Values "


        #商品主信息其实行
        mainInfoBeginRow = beginRow + 1
        # print("mainInfoBeginRow: "+ str(mainInfoBeginRow))
        #信息分界行
        separatedRow = personal_max_row(actualSheet,"出库类型")

        tmpPoint = separatedRow-1
        tmpInfo = str(actualSheet.cell(row=tmpPoint,column=1).value)

        #获取主信息最后一行位置
        while tmpInfo is None or tmpInfo == "" or tmpInfo == "None" or "合计" in tmpInfo:
            tmpPoint = tmpPoint - 1#默认商品信息最后一行和“出库类型信息”这一行只相差一行
            tmpInfo = str(actualSheet.cell(row=tmpPoint, column=1).value)
            if "合计装箱数" in tmpInfo:
                allNumInBox = str(actualSheet.cell(row=tmpPoint, column=2).value)

        # allNumInBox = ""
        # if "合计装箱数" in str(actualSheet.cell(row=tmpPoint-1,column=1).value):
        #     allNumInBox = actualSheet.cell(row=tmpPoint-1,column=2).value
        #     tmpPoint = tmpPoint - 1
        #
        # if "合计" in str(actualSheet.cell(row=tmpPoint-1,column=1).value):
        #     tmpPoint = tmpPoint - 1

        lastInfoRow = tmpPoint
        # print("last info row : " + str(lastInfoRow))

        #获取所有主信息
        while mainInfoBeginRow <= lastInfoRow:
            tmpColPos = 1
            tmpSql = ""
            while tmpColPos <= maxCol:
                tmpSql = tmpSql + "'" + str(actualSheet.cell(row=mainInfoBeginRow,column = tmpColPos).value)+ "',"
                tmpColPos = tmpColPos + 1

            detailSql = detailSql + \
                        "('" + ''.join(str(uuid.uuid4()).split('-')) + "','"+ \
                        stores_sheet_id + "','"+ \
                        orders_id + "',"+ \
                        tmpSql[:-1] + "),"

            mainInfoBeginRow = mainInfoBeginRow + 1

        detailSql = detailSql[:-1]
        print(detailSql)

        ######获取其他内容（非商品内容信息）
        # print("获取其他内容（非商品内容信息）")
        maxRow = personal_max_row(actualSheet, "要求到达时间")
        otherInfoColKey = 1
        otherInfoColValue = 2
        tmpRow = separatedRow
        storesSheetValueSql = ""
        while tmpRow <= maxRow:
            storesSheetSql = storesSheetSql + str(getOtherValue(actualSheet.cell(row=tmpRow, column=otherInfoColKey).value))+","
            storesSheetValueSql = storesSheetValueSql + "'" + str(actualSheet.cell(row=tmpRow, column=otherInfoColValue).value)+"',"
            tmpRow = tmpRow + 1

        storesSheetSql = storesSheetSql[:-1]
        storesSheetSql = storesSheetSql + ") Values ('" + stores_sheet_id +"','" + orders_id + "','"+ allNumInBox + "'," + storesSheetValueSql[:-1] + ")"
        print(storesSheetSql)


        ######db example####
        #获取sql语句拼装信息
        # try:
        #     affectRows = cursor.execute(storesSheetSql)
        #     # print(affectRows)
        #     print(detailSql)
        #     affectRows = cursor.execute(detailSql)
        #     # print(affectRows)
        #     result = cursor.fetchall()
        #     # print(result) #result 类型： 多元类型，(x1,x2,...) | xn {n = 查询数据的数量， len(x) = 每条数据的所有字段值}
        #     db.commit()
        #
        # except:
        #     print("===="+ readFilepath + ":" + sheetName + "====")


        print("++++++++++++++++++"+ readFilepath + ":" + sheetName  +"++++++++++++++++++++")

db = closeDB(db)