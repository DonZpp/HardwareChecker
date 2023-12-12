import openpyxl
import GetConfig


WORK_DIR = GetConfig.WORK_DIR
CONFIG_FILE = GetConfig.CONFIG_FILE
CONFIG_SHEET = GetConfig.CONFIG_SHEET

NUM_COL_TYPE = GetConfig.NUM_COL_TYPE
NUM_COL_KEY = GetConfig.NUM_COL_KEY
NUM_COL_VAL = GetConfig.NUM_COL_VAL
LIST_ITEM_NUM = GetConfig.NUM_COL_LIST_ITEM_NUM
ITEM_RECORDS_NUM = GetConfig.NUM_COL_LIST_ITEM_RECORDS_NUM

TYPE_LIST = GetConfig.TYPE_LIST
TYPE_LIST_ITEM = GetConfig.TYPE_LIST_ITEM
TYPE_RECORDS = GetConfig.TYPE_RECORDS


def ReadListItem(lst : list, ws : openpyxl.worksheet.worksheet, row, nItemRec):
    i = 0
    dic = {}
    while i < nItemRec:
        strKey = ws.cell(row = row, column = NUM_COL_KEY).value
        strVal = ws.cell(row = row, column = NUM_COL_VAL).value
        dic[strKey] = strVal
        i = i + 1
        row = row + 1
    lst.append(dic)
    return row

def ReadList(dicConf : dict, ws : openpyxl.worksheet.worksheet, row):
    strListName = ws.cell(row = row, column = NUM_COL_KEY).value
    dicConf[strListName] = list()
    nItemNum = int(ws.cell(row = row, column = LIST_ITEM_NUM).value)
    nItemRec = int(ws.cell(row = row, column = ITEM_RECORDS_NUM).value)
    i = 0
    row = row + 1
    while i < nItemNum:
        row = ReadListItem(dicConf[strListName], ws, row, nItemRec)
        i = i + 1
    return row

def ReadRecords(dicConf : dict, ws : openpyxl.worksheet.worksheet, row):
    strKey = ws.cell(row=row, column=NUM_COL_KEY).value
    strVal = ws.cell(row=row, column=NUM_COL_VAL).value
    dicConf[strKey] = strVal
    return row + 1


ReadObjFunc = {
    TYPE_LIST : ReadList,
    TYPE_RECORDS : ReadRecords
}



def ReadConfigExcel():
    wb = openpyxl.open(WORK_DIR + CONFIG_FILE)
    ws = wb[CONFIG_SHEET]
    i = 1

    dicConf = {}
    while ws.cell(row = i, column=1).value != None:
        strType = ws.cell(row = i, column=NUM_COL_TYPE).value
        i = ReadObjFunc[strType](dicConf, ws, i)
    return dicConf


def CheckListItem(dicItemLocal, dicItemModel):
    for k, v in dicItemLocal.items():
        if type(v) != type(list()):
            if v != dicItemModel[k]:
                return False
    return True
            
       

def CheckList(lstLocal, lstModel, strKey):
    setEq = list()
    for itemLocal in lstLocal:
        for itemModel in lstModel:
            if CheckListItem(itemLocal, itemModel) :
                setEq.append(itemLocal)
                break
    if len(setEq) != len(lstLocal):
        print('sth error!', strKey, 'Local:', lstLocal, 'Model:', lstModel)
        raise


def Check():
    dicModel = ReadConfigExcel()
    dicLocal = GetConfig.GetConfig()

    for k, v in dicLocal.items() :
        if type(v) != type(list()):
            if v != dicModel[k]:
                print('sth error!', k, 'Local:', dicLocal[k], 'Model:', dicModel[k])
                raise
        else:
            if(type(dicModel[k]) != type(list())):
                print('sth error!', k, 'Local:', dicLocal[k], 'Model:', dicModel[k])
                raise
            else:
                CheckList(v, dicModel[k], k)

    print('Good!!!!!!!!!!!!!!')


Check()