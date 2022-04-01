import re
import os


path="/Users/praette/Desktop/2960+/RawInventory_Export_2022-Mar-10_07-12-36/Config/"

flags={"AAA":False,"Stack":False,"STP":False,"PO":False,"DHCP":False,"DNS":False,"Qos":False,"SNMP":False,"PS":False,"Banner":False,"Http":False}

for fil in os.listdir(path):
	os.chdir(r'/Users/praette/Desktop/2960+/RawInventory_Export_2022-Mar-10_07-12-36/Config/')
	f=open(fil, 'r')
	#print(f.name)
	data=f.readlines()
	for line in data:
		if re.search("^aaa\s.*",line):
			flags["AAA"]=True
		if re.search("^switch\s\dprovision",line):
			flags["Stack"]=True
		if re.search("^spanning-tree\s.*",line):
			flags["STP"]=True
		if re.search("^interface\sport-channel.*",line):
			flags["PO"]=True
		if re.search("^ip\sdhcp\spool.*",line):
			flags["DHCP"]=True
		if re.search("^ip\sname-server.*",line):
			flags["DNS"]=True
		if re.search("^mls\sqos.*",line):
			flags["Qos"]=True
		if re.search("^snmp-server.*",line):
			flags["SNMP"]=True
		if re.search("^\sswitchport\sport-security.*",line):
			flags["PS"]=True
		if re.search("^banner\smotd.*",line):
			flags["Banner"]=True
		if re.search("^ip\shttp.*",line):
			flags["Http"]=True
			print(f.name)
		if re.search("^errdisable\srecovery.*",line):
			print(line)
print(flags)

#result
#{'AAA': True, 'Stack': False, 'STP': True, 'PO': False, 'DHCP': True, 'DNS': True, 'Qos': True, 'SNMP': True, 'PS': True, 'Banner': True, 'Http': True}
