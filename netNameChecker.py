import os
import xlwings as xw



os.chdir("/Users/arvid/Desktop")
wb = xw.Book("DOT Circuits tracking.xls")
sheet = wb.sheets("Circuit & IP info")


def findDescription():

	counter = 0
	for line in siteFile:
		#print(line)
		#print(counter)

		counter += 1
		if "Description" in line and counter < 4:
			interfaceDescription = line.split(" ")
			for words in interfaceDescription:
				if "Netname" in words:
		#			print(words)
					combo.insert(1,words)
					counter = 0
		#			print(combo)
				if "Circuit" in words:
					#print (words)
					combo.insert(2,words)
					counter += 0
		#			print(combo)
					return
		if (counter >= 4 ):
			combo[:] = []
		#	print(combo)
			return

def lookUpName(routerName):
	for i in range(2,97):
		cellValue =sheet.range("A" + str(i)).value 
		#print(cellValue)
		routerName = routerName.replace("-","")
		#print(routerName)
		cellValue = cellValue.replace(" ","")
		if routerName == cellValue:
			print (routerName + "  " + cellValue)
			print (i)
			matchCiruit(i,Circuit)
			return 
		else:
			pass

def matchCiruit(i,Circuit):
	#print(i)
	#print(Netname)
	execlCiruitNameCell = sheet.range("D" + str(i))
	execlCiruitName = execlCiruitNameCell.value.strip()
	print (execlCiruitName)
	if "-" in Circuit:
		Circuit = Circuit.split("-")[1].strip()
	print(Circuit)
	if execlCiruitName == Circuit:
		print("its a match")
		execlCiruitNameCell.color = (0,255,0)
	else:
		execlCiruitNameCell.color = (255,0,0)
	checkNetName(i,Netname)



def checkNetName(i,Netname):
	execlNetNameCell = sheet.range("B" + str(i))
	execlNetName = execlNetNameCell.value.strip()
	Netname = Netname.split("-")[1]
	print(Netname)	

	if "-" in execlNetName:
		execlNetName = execlNetName.replace("-","")
		print(execlNetName)

	if execlNetName == Netname:
		execlNetNameCell.color = (0,255,0)
	else:
		execlNetNameCell.color = (255,0,0)
	

	


siteFile = open("390 Kent Ave rt", "r")
combo = []
for line in siteFile:
	#find name of rtr
	if "#" in line:
		#print(line)
		routerName = line.split("#")[0].split("_")[0]
		#print(routerName)
	if "GigabitEthernet" in line:
		#print(line.split(" "))
		gigabit = line.split(" ")[0]
		combo.insert(0,gigabit)
		#print(gigabit)
		#print(combo)
		findDescription()
		
	elif "Serial" in line:
		#print(line.split(" "))
		serial = line.split(" ")[0]
		combo.insert(0,serial)
		#print(serial)
		findDescription()

	if len(combo) == 3:
		print (combo)
		break

interface = combo[0]
Netname = combo[1]
Circuit = combo[2]

lookUpName(routerName)




