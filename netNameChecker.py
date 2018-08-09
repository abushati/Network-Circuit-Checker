import os
import xlwings as xw
import re




os.chdir("/Users/arvid/Desktop/Router Information")
wb = xw.Book("/Users/arvid/Desktop/DOT Circuits tracking.xls")
sheet = wb.sheets("Circuit & IP info")


def findDescription(siteFile,combo):

	counter = 0
	for line in siteFile:
		#print(line)
		
		counter += 1
		#print(counter)
		#print (line)
		if counter < 5 and "  Description:" in line:
			#print("its in a line")
			interfaceDescription = line.split(" ")

			if "Netname" in interfaceDescription:
				#print(interfaceDescription[interfaceDescription.index("Netname") + 1])
				
				networkName = interfaceDescription[interfaceDescription.index("Netname") + 1] 
				networkName = networkName.replace("*","")

				combo.insert(1,networkName)
				counter = 0
				#print(combo)
			if "Circuit" in interfaceDescription:
				circuitName = interfaceDescription[interfaceDescription.index("Circuit") + 1] 
				#print (words)
				if "#" or "*" in circuitName:
					circuitName = circuitName.replace("*","").replace("#","")

				combo.insert(2,circuitName)
				counter = 0
				print(combo)
				return
		if (counter >= 5 ):
			combo[:] = []
			print(combo)
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
	

listOfFiles = os.listdir("/Users/arvid/Desktop/Router Information")



combo = []
for routerFile in listOfFiles:
	print (routerFile)

	siteFile = open(routerFile, "r")
	if routerFile.endswith(".txt"):

	#routerName = siteFile.readlines()[-1]
	#print(routerName)
		for line in siteFile:

			#find name of rtr
			
				#print(routerName)
			if "GigabitEthernet" in line:
				#print(line.split(" "))
				gigabit = line.split(" ")[0]
				combo.insert(0,gigabit)
				#print(gigabit)
				print(combo)
				findDescription(siteFile,combo)
				
			elif "Serial" in line:
				#print(line.split(" "))
				serial = line.split(" ")[0]
				combo.insert(0,serial)
				#print(serial)
				print(combo)
				findDescription(siteFile,combo)
				

			if len(combo) == 3:
				interface = combo[0]
				Netname = combo[1]
				Circuit = combo[2]
				print(interface)
				print(Netname)	
							
				print (combo)
				break
			else:
				continue



		lookUpName(routerName)




