import psutil
import platform
import GPUtil
import wmi
import os
import subprocess


WORK_DIR = 'Z:\\Project\\HardwareConfigurationCheck\\'
CONFIG_FILE = 'ModelConfig.xlsx'
CONFIG_SHEET = 'Conf'

TYPE_LIST = 'List'
TYPE_LIST_ITEM = 'List Item'
TYPE_RECORDS = 'Record'

NUM_COL_TYPE = 1
NUM_COL_KEY = 2
NUM_COL_VAL = 3
NUM_COL_LIST_ITEM_NUM = 3
NUM_COL_LIST_ITEM_RECORDS_NUM = 4


def GenModel(*args):
    def Model():
        infos = {}
        for key in args:
            infos[key] = ''
        return infos
    return Model


#Check Item Names
#SYSTEM
SYSTEM_NAME = 'System Name'
SYSTEM_VERSION = 'System Version'
SYSTEM_MACHINE = 'System Machine'
#CPU
CPU_NAME = "CPU Name"
CPU_PHYSIC_CORES = 'CPU Physiccal Cores'
CPU_TOTAL_CORES = 'CPU Total Cores'
CPU_MAX_FREQ = 'CPU Max Frequency'
#GPU
GPU_INFO = 'GPU Info' # dictionary to check more than one gpu
GPU_NAME = 'GPU Name'
GPU_MEMO = 'GPU Memory'
#MEMORY
MEMO_TOTAL_SIZE = "Memory Size"
#DISK
DISK_INFO = 'Disk Info' # list to check more than one disk
DISK_SIZE = "Disk Size"
DISK_TYPE = "Disk Type"  # HDD / SSD


CheckItem = {}

#Add System Check Items
CheckItem[SYSTEM_NAME] = ''
CheckItem[SYSTEM_VERSION] = ''
CheckItem[SYSTEM_MACHINE] = ''

#Add CPU Check Items
CheckItem[CPU_NAME] = ''
CheckItem[CPU_PHYSIC_CORES] = ''
CheckItem[CPU_TOTAL_CORES] = ''
CheckItem[CPU_MAX_FREQ] = ''

#Add GPU Check Items
GPUInfoModel = GenModel(GPU_NAME, GPU_MEMO)
CheckItem[GPU_INFO] = list()

#Add Memory Check Items
CheckItem[MEMO_TOTAL_SIZE] = ''

#Add Disk Check Items
DiskInfoModel = GenModel(DISK_SIZE, DISK_TYPE)
CheckItem[DISK_INFO] = list()


def GetConfig():
    dicConf = {}
    # get system info
    uname = platform.uname()
    dicConf[SYSTEM_NAME] = str(uname.system)
    dicConf[SYSTEM_VERSION] = str(uname.version)
    dicConf[SYSTEM_MACHINE] = str(uname.machine)

    # get cpu info
    for p in wmi.WMI().Win32_Processor():
        processor = p

    cpufreq = psutil.cpu_freq()

    dicConf[CPU_NAME] = str(processor.Name)
    dicConf[CPU_PHYSIC_CORES] = str(psutil.cpu_count(logical=False))
    dicConf[CPU_TOTAL_CORES] = str(psutil.cpu_count(logical=True))
    dicConf[CPU_MAX_FREQ] = str(cpufreq.max)

    # get GPU info
    gpus = GPUtil.getGPUs()
    dicConf[GPU_INFO] = list()
    for gpu in gpus:
        InfoModel = GPUInfoModel()
        InfoModel[GPU_NAME] = str(gpu.name)
        InfoModel[GPU_MEMO] = str(gpu.memoryTotal)
        dicConf[GPU_INFO].append(InfoModel)

    # get Memory info
    svmem = psutil.virtual_memory()
    dicConf[MEMO_TOTAL_SIZE] = str(svmem.total)

    # get Disk info
    dicConf[DISK_INFO] = list()
    res = subprocess.run('powershell Get-PhysicalDisk | Select MediaType, Size', capture_output = True)
    lst = res.stdout.__str__()[2:-1].split('\\r\\n')
    lstRemove = list()
    for s in lst:
        if (s.find('SSD') == -1 and s.find('HDD') == -1):
            lstRemove.append(s)
    for strRemv in lstRemove:
        lst.remove(strRemv)
    for strRecords in lst:
        InfoModel = DiskInfoModel()
        lstRecords = strRecords.split()
        InfoModel[DISK_TYPE] = str(lstRecords[0])
        InfoModel[DISK_SIZE] = str(lstRecords[1])
        dicConf[DISK_INFO].append(InfoModel)
    return dicConf


#test
def PrintConfigInfo():
    for k, v in GetConfig().items():
        print(k, '|', v)

