"""
Version1.01 date:2020-7-6
Author��Louisyang
"""
import os   # ����ϵͳ������
import xlrd # ����Excel������
import re   # ����������ʽ��

# �����Excel����
def read(row_value, column_value): 
    # ��ȡĳ�У�ĳ�е�Ԫ������
    string = worksheet.cell_value(row_value, column_value)
    return string

# ��xlrd���ex1.xlsx�ļ�
workbook = xlrd.open_workbook(filename='ex1.xlsx')
# ��1�Ź�����
worksheet = workbook.sheets()[0] 

# ����ĸ��ת��Ϊ������
words = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
def col_num(col_0):
    col_0 = col_0.upper()
    col_0 = words.index(col_0)
    return col_0

# �������д�0��ʼ��
def row_num(row):
    row = row - 1
    return row

# ����DBC_Header����
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

# ����һ�����б�����Header
value_table = []
value_table.append(l)
value_table.append("\n")

for i in range ((row_num(4)),worksheet.nrows): 
    #ȡ��i�е��ַ�������
    row = worksheet.row_values(i) 
    if row[col_num("J")] == "":
        value_table.append("\n")
        id = row[4]
        id = id.replace("0x", "")
        s = int (id, 16)
    #�ο�BO_ 796 EPS3: 3 EPS /////BO_ ID messageNAME:length��Message_Node
        value_table.append(f'BO_ {s} {row[1]}: {int(row[col_num("H")])} {row[0]}')
        value_table.append("\n")
        continue
    else:
    #�ο�SG_ GW_ESC_ESCOFF : 0|1@0+ (1,0) [0|1] ""  Radar /////ID��messageNAME:length��ECU
        node = row[col_num("A")]
        if node == "ADASC":
            ecu = "Vector__XXX"
        else:
            ecu = "ADASC"
        #�ο�SG_ message_name : start_bit|length @0+ (factor,offset) [min|max] ""  Rx_Node
        value_table.append(f' SG_ {row[col_num("J")]} : {int(row[col_num("N")])}|{int(row[col_num("L")])}@0+ (1,0) [0|{2**int(row[col_num("L")])-1}] "" {ecu}')
        value_table.append("\n")
        
value_table.append("\n"*2)


c = """
BA_DEF_ BO_  "GenMsgCycleTime" INT 0 50000;
BA_DEF_DEF_  "GenMsgCycleTime" 0;

"""
value_table.append(c)

for i in range ((row_num(4)),worksheet.nrows): 
    #ȡ��i�е��ַ�������
    row = worksheet.row_values(i) 
    # ����Ϊ�յ�һ�б��
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
            # ����Message���ź�����
            value_table.append(f'BA_ "GenMsgCycleTime" BO_ {s} {int(cycle)};')
            value_table.append("\n") 
            
value_table.append("\n"*2)

# ���嵥Ԫ����ϴ����
def clean(val): 
    val = re.sub(r'[\n]', '', val) 
    val = re.sub(r'[\xa0]', '', val) 
    #����ƥ�䣬��@�滻�����֣�������
    val = re.sub(r'[0-9][:��]', '@ ', val) 
    return val

# ����value_descriptionƴ�Ӻ���
def transfer(unit): 
    unit = " " + unit
    # ��@�и�ָ��ַ���Ϊ�б�
    unit = re.split("@ ", unit)
    # ɾ����һ��Ԫ��
    del unit[0]
    msg = ""
    # ƴ�����ֺ�ֵ
    for d in range(len(unit)): 
        # VAL_ 1043 Mp5_LSS_main_SW 1 "LSS ON" 0 "LSS OFF" ;
        # l = f'VAL_ {s} {row[col_num("J")]} {val_0};'
        msg = msg + str(d) +" "+ "\"" + unit[d] + "\"" + " "
        """
        msg = msg + f' {str(d)} "{unit[d]}" '
        """
    return msg

for i in range ((row_num(4)),worksheet.nrows): 
    #ȡ��i�е��ַ�������
    row = worksheet.row_values(i) 
    if row[col_num("J")] == "":
        continue
    else:
        id = row[col_num("E")]
        # ʮ������תʮ����
        id = id.replace("0x", "")
        s = int (id, 16)
        val_0 = row[col_num("X")]
        val_0 = clean(val_0)
        
        #��õ�i�е�value_description
        val_0 = transfer(val_0) 
        
        # VAL_ 1043 Mp5_LSS_main_SW 1 "LSS ON" 0 "LSS OFF" ;
        # ����value_table����
        # VAL_ ID Message_name order value order value����
        l = f'VAL_ {s} {row[col_num("J")]} {val_0};'
        value_table.append(l)
        value_table.append("\n")
             
with open ('ex0.dbc', 'w+', encoding = 'GB2312') as f:
    f.writelines(value_table) 

f = open('ex0.dbc', 'rt')
s =f.read()
print(s)
f.close()
