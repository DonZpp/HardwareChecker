import subprocess
res = subprocess.run('powershell Get-PhysicalDisk | Select MediaType, Size', capture_output = True)
lst = res.stdout.__str__()[2:-1].split('\\r\\n')
lstRemove = list()
for s in lst:
    if (s.find('SSD') == -1 and s.find('HDD') == -1):
        lstRemove.append(s)
for strRemv in lstRemove:
    lst.remove(strRemv)