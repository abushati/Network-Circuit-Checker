import os
import xlwings as xw
import time
from difflib import SequenceMatcher


os.chdir("/Users/arvid/Desktop/Router Information")
wb = xw.Book("/Users/arvid/Desktop/DOT Circuits tracking.xls")
sheet = wb.sheets("Circuit & IP info")

SiteNameCombos = {}
sitesNotMatched = []


def findDescription(siteFile,combo):

	lineCounter = 0
	for line in siteFile:
		#print(line)
		
		lineCounter += 1
		#print(lineCounter)
		#print (line)
		if lineCounter < 5 and "  Description:" in line:
			#print("its in a line")
			
			interfaceDescription = line.replace("Netname-","Netname ").replace("Netname:","Netname ")\
			.replace("Netname-","Netname ").replace("Netname:","Netname ").replace("Circuit-","Circuit ")\
			.replace("Circuit:","Circuit ").replace("***Netname","Netname").replace("***netname","Netname")\
			.replace("***Circuit","Circuit").replace("***circuit#","Circuit").replace("circuit-","circuit ")\
			.replace("Circuit#","Circuit").replace("Netname#","Netname").replace("**Circuit","Circuit").split(" ")
			if " " in interfaceDescription:
				combo.remove(" ")

			#BIG DEBUGGER!!!!!!!!!!	
			#print(interfaceDescription)

			if "Netname" in interfaceDescription:
				#print(interfaceDescription[interfaceDescription.index("Netname") + 1])
				networkName = interfaceDescription[interfaceDescription.index("Netname") + 1] 
				networkName = networkName.replace("*","").replace("#","").replace("-","")
				combo.insert(1,networkName)
				lineCounter = 0
				#print(combo)
			if ("Circuit") in interfaceDescription:
				circuitName = interfaceDescription[interfaceDescription.index("Circuit") + 1] 
				#print (words)
				#if "#" or "*" or "-"in circuitName:
				circuitName = circuitName.replace("*","").replace("#","").replace("-","")
				combo.insert(2,circuitName)
				lineCounter = 0
				#print(combo)
				return
			if ("circuit") in interfaceDescription:
				circuitName = interfaceDescription[interfaceDescription.index("circuit") + 1] 
				#print (words)
				#if "#" or "*" or "-"in circuitName:
				circuitName = circuitName.replace("*","").replace("#","").replace("-","")
				combo.insert(2,circuitName)
				lineCounter = 0
				return	
			elif ""	:
				pass	
		
		#the description is always 3 lines below the start of the file,
		#if the line counter passes 5, the file doesnt contain the info
		if (lineCounter >= 5 ):
			combo[:] = []
			#print(combo)
			return
		else: 
			continue
		


def lookUpName(routerName):
	global i
	numOfMatches = 0
	for i in range(2,97):
		cellValue = sheet.range("A" + str(i)).value 
		cellValue= cellValue.strip().replace(" ","")

		if cellValue != "None":
			#print(cellValue)

			ratioOfWord = SequenceMatcher("",routerName, cellValue).ratio()
			print(numOfMatches)
			if ratioOfWord > .80:
				print(routerName + "-->" + cellValue)
				print(ratioOfWord)
				#time.sleep(6)
				numOfMatches += 1 
				print(numOfMatches)
				routersMatched.append(routerName)
				#SiteNameCombos.add({"CRT site name" : routerName , "Excel site name" : cellValue})
				#print( SiteNameCombos)
				#time.sleep(10)
				return

	
	#for sites in sitesNotMatched :
	#	print(sites)


def matchCiruit(i,Circuit):
	#print(i)
	#print(Netname)
	execlCiruitNameCell = sheet.range("D" + str(i))
	execlCiruitName = execlCiruitNameCell.value.strip()
	
	
	#print("In the router: " + Circuit)
	#time.sleep(2)
	if execlCiruitName == Circuit:
		#print("its a match")
		execlCiruitNameCell.color = (0,255,0)
	else:
		execlCiruitNameCell.color = (255,0,0)
	



def checkNetName(i,Netname):
	execlNetNameCell = sheet.range("B" + str(i))
	execlNetName = execlNetNameCell.value.strip()
	#print(Netname)	

	if "-" in execlNetName:
		execlNetName = execlNetName.replace("-","")
		#print(execlNetName)

	if execlNetName == Netname:
		execlNetNameCell.color = (0,255,0)
	else:
		execlNetNameCell.color = (255,0,0)
	

listOfFiles = os.listdir("/Users/arvid/Desktop/Router Information")

routerOnCRT = []
routersMatched = []
combo = []
for routerFile in listOfFiles:
	
	if routerFile.endswith(".txt"):
		siteFile = open(routerFile, "r")

		#print (routerFile)

		for line in siteFile:
			
				#find name of rtr
				#print(routerName)
			if "GigabitEthernet" in line:
				#print(line.split(" "))
				gigabit = line.split(" ")[0]
				combo.insert(0,gigabit)
				#print(gigabit)
				#print(combo)
				findDescription(siteFile,combo)
				
			elif "Serial" in line:
				#print(line.split(" "))
				serial = line.split(" ")[0]
				combo.insert(0,serial)
				#print(serial)
				#print(combo)
				findDescription(siteFile,combo)
				

			if len(combo) == 3:
				#print('this file is good ' + routerFile)
				#print(combo )
				try:
					#os.rename("/Users/arvid/Desktop/Router Information/" + routerFile, "/Users/arvid/Desktop/updatedRI/" + routerFile)
					routerName = siteFile.readlines()[-1].strip(".txt")
					routerName = routerName.replace("-","").replace("_","").strip()
					#routerName = routerName.split(" ")
					print("Name on router is : " + routerName)	
					#time.sleep(4)
					#print(combo)
					routerOnCRT.append(routerName)
					interface = combo[0]
					Netname = combo[1].strip()
					Circuit = combo[2].strip()
					combo[:] =[]
					#print(interface)
					#print(Netname)	
					#print(Circuit)		
					#print (combo)

					lookUpName(routerName)


					matchCiruit(i, Circuit)
					checkNetName(i,Netname)
					break
				except AttributeError:
					continue

print(routerOnCRT)
print(routersMatched)






