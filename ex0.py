"""
Version1.01 date:2020-7-6
Author：Louisyang
"""
import os   # 导入系统操作库
import xlrd # 导入Excel操作库
import re   # 导入正则表达式库

# 定义读Excel函数
def read(row_value, column_value): 
    # 读取某行，某列单元格内容
    string = worksheet.cell_value(row_value, column_value)
    return string

# 调xlrd库打开ex1.xlsx文件
workbook = xlrd.open_workbook(filename='ex1.xlsx')
# 打开1号工作表
worksheet = workbook.sheets()[0] 

# 把字母列转化为数字列
words = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
def col_num(col_0):
    col_0 = col_0.upper()
    col_0 = words.index(col_0)
    return col_0

# 把数字行从0开始数
def row_num(row):
    row = row - 1
    return row

# 补充DBC_Header内容
l = """
VERSION ""

NS_ :
	BA_
	BA_DEF_
	BA_DEF_DEF_
	BA_DEF_DEF_REL_
	BA_DEF_REL_
	BA_DEF_SGTYPE_
	BA_REL_
	BA_SGTYPE_
	BO_TX_BU_
	BU_BO_REL_
	BU_EV_REL_
	BU_SG_REL_
	CAT_
	CAT_DEF_
	CM_
	ENVVAR_DATA_
	EV_DATA_
	FILTER
	NS_DESC_
	SGTYPE_
	SGTYPE_VAL_
	SG_MUL_VAL_
	SIGTYPE_VALTYPE_
	SIG_GROUP_
	SIG_TYPE_REF_
	SIG_VALTYPE_
	VAL_
	VAL_TABLE_

BS_:

BU_: Radar ESC SAS ACU GW Tester RADAR Camera EPS TCU ADASC CGW
"""

# 定义一个空列表，补充Header
value_table = []
value_table.append(l)
value_table.append("\n")

for i in range ((row_num(4)),worksheet.nrows): 
    #取第i行的字符串内容
    row = worksheet.row_values(i) 
    if row[col_num("J")] == "":
        value_table.append("\n")
        id = row[4]
        id = id.replace("0x", "")
        s = int (id, 16)
    #参考BO_ 796 EPS3: 3 EPS /////BO_ ID messageNAME:length、Message_Node
        value_table.append(f'BO_ {s} {row[1]}: {int(row[col_num("H")])} {row[0]}')
        value_table.append("\n")
        continue
    else:
    #参考SG_ GW_ESC_ESCOFF : 0|1@0+ (1,0) [0|1] ""  Radar /////ID、messageNAME:length、ECU
        node = row[col_num("A")]
        if node == "ADASC":
            ecu = "Vector__XXX"
        else:
            ecu = "ADASC"
        #参考SG_ message_name : start_bit|length @0+ (factor,offset) [min|max] ""  Rx_Node
        value_table.append(f' SG_ {row[col_num("J")]} : {int(row[col_num("N")])}|{int(row[col_num("L")])}@0+ (1,0) [0|{2**int(row[col_num("L")])-1}] "" {ecu}')
        value_table.append("\n")
        
value_table.append("\n"*2)


c = """
BA_DEF_ BO_  "GenMsgCycleTime" INT 0 50000;
BA_DEF_DEF_  "GenMsgCycleTime" 0;

"""
value_table.append(c)

for i in range ((row_num(4)),worksheet.nrows): 
    #取第i行的字符串内容
    row = worksheet.row_values(i) 
    # 跳过为空的一行表格
    if row[col_num("A")] == "":
        continue
    else:
        cycle = row[col_num("G")]
        if cycle == "":
            continue
        else:
            id = row[4]
            id = id.replace("0x", "")
            s = int (id, 16)
            # BA_ "GenMsgCycleTime" BO_ 777 100;
            # 定义Message的信号周期
            value_table.append(f'BA_ "GenMsgCycleTime" BO_ {s} {int(cycle)};')
            value_table.append("\n") 
            
value_table.append("\n"*2)

# 定义单元格清洗函数
def clean(val): 
    val = re.sub(r'[\n]', '', val) 
    val = re.sub(r'[\xa0]', '', val) 
    #正则匹配，用@替换“数字：”内容
    val = re.sub(r'[0-9][:：]', '@ ', val) 
    return val

# 定义value_description拼接函数
def transfer(unit): 
    unit = " " + unit
    # 用@切割、分割字符串为列表；
    unit = re.split("@ ", unit)
    # 删除第一个元素
    del unit[0]
    msg = ""
    # 拼接数字和值
    for d in range(len(unit)): 
        # VAL_ 1043 Mp5_LSS_main_SW 1 "LSS ON" 0 "LSS OFF" ;
        # l = f'VAL_ {s} {row[col_num("J")]} {val_0};'
        msg = msg + str(d) +" "+ "\"" + unit[d] + "\"" + " "
        """
        msg = msg + f' {str(d)} "{unit[d]}" '
        """
    return msg

for i in range ((row_num(4)),worksheet.nrows): 
    #取第i行的字符串内容
    row = worksheet.row_values(i) 
    if row[col_num("J")] == "":
        continue
    else:
        id = row[col_num("E")]
        # 十六进制转十进制
        id = id.replace("0x", "")
        s = int (id, 16)
        val_0 = row[col_num("X")]
        val_0 = clean(val_0)
        
        #获得第i行的value_description
        val_0 = transfer(val_0) 
        
        # VAL_ 1043 Mp5_LSS_main_SW 1 "LSS ON" 0 "LSS OFF" ;
        # 定义value_table内容
        # VAL_ ID Message_name order value order value……
        l = f'VAL_ {s} {row[col_num("J")]} {val_0};'
        value_table.append(l)
        value_table.append("\n")
             
with open ('ex0.dbc', 'w+', encoding = 'GB2312') as f:
    f.writelines(value_table) 

f = open('ex0.dbc', 'rt')
s =f.read()
print(s)
f.close()
