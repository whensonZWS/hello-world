import pandas as pd
import sys
import configparser as cfp
import math
# new python
# read in arguments, test for ending file name
# only allow 2 file, 1 xlsx file and 1 map/yrm file

# scan the map file, and delete all the ISS triggers
# generate triggers from xlsx file
# insert all the ISS triggers into the map file



if len(sys.argv)!=3:
	print("This script requires exactly 2 input files!")
	#eeeee=input()
	sys.exit()

temp1=sys.argv[1][-4:len(sys.argv[1])]
temp2=sys.argv[2][-4:len(sys.argv[2])]
if temp1==".xls" or temp1=="xlsx":
	tableName=sys.argv[1]
	mapName=sys.argv[2]
elif temp2==".xls" or temp2=="xlsx":
	tableName=sys.argv[2]
	mapName=sys.argv[1]
else:
	print("This script requires exactly 1 excel file (*.xls or *.xlsx)!")
print("table name =",tableName)
print("map name =",mapName)



#generating triggers, td: trigger data
t=pd.read_excel(tableName,converters={'Teams':str})
n=len(t.Time)
	
#id generator for tag and triggers
tagID = lambda i: "A" + str(1550000+i)
triggerID = lambda i: "B" + str(1550000+i)
#convert waypoint from number into alphabet format (weird ini requirement)
def wp(i):
	alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	first=i//26-1
	second=i%26
	result=""
	if first>=0:
		result=result+alphabet[first]
	result=result+alphabet[second]
	return result
	
tag=[]
trigger=[]
events=[]
actions=[]

for i in range(0,n):
	tag.append("0,{},{}".format(t.Name[i],triggerID(i)))
	trigger.append("UnitedStates,<none>,{},0,1,1,1,0".format(t.Name[i]))
	events.append("1,13,0,{}".format(t.Time[i]))
	waypoints=list(map(int,t.WayPoint[i].split(',')))
	sounds=[]
	text=[]
	if str(t.Sound[i])!="nan":
		sounds=t.Sound[i].split(',')
	if str(t.Text[i])!="nan":
		text=t.Text[i].split(',')
	#text will be put into actions
	ac=str(len(waypoints)+len(sounds)+len(text))
	for wp_num in waypoints:
		ac=ac+",80,1,{},0,0,0,0,{}".format(t.Teams[i],wp(wp_num))
	for sd_id in sounds:
		ac=ac+",19,7,{},0,0,0,0,A".format(sd_id)
	for txt in text:
		ac=ac+",11,4,{},0,0,0,0,A".format(txt)
	actions.append(ac)


	
# delete existing ISS trigger
#mapName=sys.argv[1]
m=cfp.ConfigParser(strict=False)
m.optionxform=str
m.read(mapName)

for key in m['Actions']:
	if str(key)[0:4]=="B155" :
		m.remove_option("Actions",key)
for key in m['Events']:
	if str(key)[0:4]=="B155" :
		m.remove_option("Events",key)
for key in m['Triggers']:
	if str(key)[0:4]=="B155" :
		m.remove_option("Triggers",key)
for key in m['Tags']:
	if str(key)[0:4]=="A155" :
		m.remove_option("Tags",key)


for i in range(0,n):
	m.set("Tags",tagID(i),tag[i])
	m.set("Triggers",triggerID(i),trigger[i])
	m.set("Events",triggerID(i),events[i])
	m.set("Actions",triggerID(i),actions[i])
		
		
with open(mapName, 'w') as configfile:
	m.write(configfile,space_around_delimiters=False)
