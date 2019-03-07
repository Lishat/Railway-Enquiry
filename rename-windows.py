import xlrd
import os
import time
Rail_database = xlrd.open_workbook("Railway enquiry.xlsx")
Stations = ["Secunderabad", "Kazipet", "Warangal", "Vijayawada", "Rajahmundry", "Visakhapatnam"]
Secunderabad = Rail_database.sheet_by_name("Secunderabad")
Kazipet = Rail_database.sheet_by_name("Kazipet")
Warangal = Rail_database.sheet_by_name("Warangal")
Vijayawada = Rail_database.sheet_by_name("Vijayawada")
Rajahmundry = Rail_database.sheet_by_name("Rajahmundry")
Visakhapatnam = Rail_database.sheet_by_name("Visakhapatnam")
name=input("Please Enter Your Name:\n")
def marquee():
	i = 0
	while(i != 50):
		os.system('cls')
		print("\n"*20)
		print(" "*i, "Hello", name)
		print("\n"*20)
		time.sleep(0.1)
		i += 1
	i = 100
	while(i != 50):
		os.system('cls')
		print("\n"*20)
		print(" "*i, "Welcome to Indian Railways")
		print("\n"*20)
		time.sleep(0.1)
		i -= 1

def load():
	os.system('cls')
	i = 1
	while i != 8:
		print("\n"*20, " "*59, "LOADING\n", " "*59, "."*i,"\n", "\n"*20)
		time.sleep(1)
		i += 1
marquee()
choice = "yes"
while(choice == "yes"):
	os.system('cls')
	print("-"*143,"\t\t\t\t\t\t\tWelcome to Rail Enquiry System", "-"*143, sep="\n")
	print("\t\t\t\t\t\tThis railway enquiry system works on the route\n", "-"*143, "\n", " "*25, "Secunderabad --> Kazipet --> Warangal --> Vijayawada --> Rajahmundry --> Visakhapatnam\n","-"*143, "\n", sep = "")
	print("1.Find Train name by Train no.")
	print("2.Find Trains between 2 stations")
	print("3.Find Schedule of a specified train by Train no.")
	print("4.Find all Trains passing through a station")
	ops = input()
	if(ops == "1"):
		train = int(input("Enter the train no.:\n"))
		load()
		os.system('cls')
		i = 0
		found = 0
		while(i != 6):
				try:
					j = 0
					while(int(eval(Stations[i]).cell(j, 0).value) != xlrd.empty_cell.value):
						if(int(eval(Stations[i]).cell(j, 0).value) == train):
							print("Train no.: ", int(eval(Stations[i]).cell(j, 0).value), "\nTrain name: ", eval(Stations[i]).cell(j, 3).value, "\nOriginating Station: ", eval(Stations[i]).cell(j, 1).value, "\nDestination Station: ", eval(Stations[i]).cell(j, 2).value)
							found = 1
							break
						j += 1
				except:
					pass
				if(found == 1):
					break
				i += 1
		if found == 0:
			print("Sorry the train doesn't run on this route")
	elif(ops == "2"):
		station = input("Enter your Boarding station\n")
		station1 = input("Enter your destination:\n")
		load()
		os.system('cls')
		print("-"*143,"\nT.no. ", "Departing Station ", " "*13, "Destination", " "*20, "Train name", " "*25, "Departure\t", " Arrival","\n", "-"*143, sep="")
		try:
			if(Stations.index(station) <= Stations.index(station1)):
				i = 0
				try:
					while(eval(station).cell(i, 0).value != xlrd.empty_cell.value):
						try:
							j  = 0
							while(eval(station1).cell(j, 0).value != xlrd.empty_cell.value):
								if(eval(station).cell(i, 0).value == eval(station1).cell(j, 0).value):
									p = len(eval(station).cell(i, 1).value)
									q = len(eval(station).cell(i, 2).value)
									r = len(eval(station).cell(i, 3).value)
									print(int(eval(station).cell(i, 0).value), eval(station).cell(i, 1).value + " "*(30 - p), eval(station).cell(i, 2).value+ " "*(30 - q), eval(station).cell(i, 3).value + " "*(35 - r), eval(station). cell(i, 4).value, eval(station). cell(i, 5).value, eval(station).cell(i, 6).value, "\t", eval(station1). cell(j, 4).value, eval(station1). cell(j, 5).value, eval(station1).cell(j, 6).value)
									break;
								j += 1
						except:
							pass
						i += 1
				except:
					pass
			else:
				print("Sorry this railway enquiry system doesn't cover this route")
		except:
			print("Sorry this railway enquiry system doesn't cover this route")
	elif(ops == "3"):
		train = int(input("Enter the train no.:\n"))
		load()
		os.system('cls')
		i = 0
		found = 0
		flag = 1
		while(i != 6):
				try:
					j = 0
					while(int(eval(Stations[i]).cell(j, 0).value) != xlrd.empty_cell.value):
						if(int(eval(Stations[i]).cell(j, 0).value) == train):
							if flag == 1:
								print("Train no.", train, "\nTrain name:", eval(Stations[i]).cell(j, 3).value)
								print("-"*50,"\n","Stations", " "*34, "Time\n","-"*50, sep="")
								flag = 0
							print(Stations[i], " "*(40 - len(Stations[i])), eval(Stations[i]).cell(j, 4).value, eval(Stations[i]).cell(j, 5).value, eval(Stations[i]).cell(j, 6).value)
							found = 1
						j += 1
				except:
					pass
				i += 1
		if found == 0:
			print("Sorry the train doesn't run on this route")
		
	elif(ops == "4"):
		station = input("Enter the name of station:\n")
		load()
		os.system('cls')
		try:
			Stations.index(station)
			try:
				print("-"*143,"\nT.no. ", "Departing Station ", " "*15, "Destination", " "*20, "Train name", " "*25, "Departure\t","\n", "-"*143, sep="")
				i = 0
				while(eval(station).cell(i, 0).value != xlrd.empty_cell.value):
					p = len(eval(station).cell(i, 1).value)
					q = len(eval(station).cell(i, 2).value)
					r = len(eval(station).cell(i, 3).value)
					print(int(eval(station).cell(i, 0).value), eval(station).cell(i, 1).value + " "*(32 - p), eval(station).cell(i, 2).value+ " "*(30 - q), eval(station).cell(i, 3).value + " "*(35 - r), eval(station). cell(i, 4).value, eval(station). cell(i, 5).value, eval(station).cell(i, 6).value)
					i += 1					
			except:
				pass
		except:
			print("This station is not covered by this railway enquiry system")
	else:
		print("Invalid choice")
	choice = input("Do you want another enquiry?\n").lower()
	os.system('cls')
