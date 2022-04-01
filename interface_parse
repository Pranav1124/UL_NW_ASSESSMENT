import re
import os


path="/Users/praette/Desktop/interview/input_folder/"

dev={}

for fil in os.listdir(path):
	f=open(fil, 'r')
	print(f.name)
	dev[os.path.basename(f.name)]={}
	int_lines=False
	int_line_CFG=""
	int_name=""
	for line in f:
	

		int_srt=re.search(r'^interface\s(.*)', line)
		#print(int_srt.group(0))
		bgp_srt=re.search(r'^router\sbgp\s\d+',line)
	
		if int_srt:
		
			int_lines=True
			int_name=int_srt.group(1)
			continue
		if int_lines:
			int_line_CFG+=line+"\n"
		if '!' in line:
			int_lines=False
		
			dev[os.path.basename(f.name)][int_name]=int_line_CFG
			int_line_CFG=""
			continue

	
	f.close()
#print(dev.keys())

	#print(dev[i])


if 'ipv6'in dev['ASBR17.txt']['Loopback0']:
	print("yes")
