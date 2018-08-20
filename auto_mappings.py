import pandas as pd
import pyautogui as pya
import time
import pyperclip


excelFilepath = input("Enter the path to your excel file (ex: /Users/jorgealfaro/Downloads/Mappings): ")
#'/Users/jorgealfaro/Downloads/yimappings2.xlsx'
df = pd.read_excel(excelFilepath, 1, usecols=[0]) #second sheet, first column
df2 = pd.read_excel(excelFilepath, 1, usecols=[1])

dflist = df['COLUMN'].tolist() # puts everything under "column" column in a list
dflist2 = df2['COLUMN2'].tolist()
print("first column contains: " + str(len(dflist)) + " cells.")
print("second column contains: " + str(len(dflist2)) + " cells.")

if (len(dflist) != len(dflist2) ):
	print("WARNING: column lengths do not match! ")

time.sleep(5) # give user time to navigate to RAD tool & place cursor on conditional 


for CODE in dflist:
	pya.click(pya.position()) #function runs through list and creates RAD conditionals

for x in range(3): pya.press('tab')

for CODE in dflist: # loops through dflist and populates conditionals according to code
	
	#pya.typewrite( "{{SubsidiaryPlanId}} = " )

	pyperclip.copy("{{SubsidiaryPlanId}} = ")  #copies code into clipboard
	pya.hotkey('command', 'v') # pastes the code 

	pya.press('tab')

	pyperclip.copy('"' + CODE + '"') #
	pya.hotkey('command', 'v')
	#pya.typewrite('"' + CODE + '"')

	if CODE != dflist[len(dflist)-1]: #if it's not the last code, navigate down to the next conditional
		for x in range(3): pya.press('tab')

print("Navigate to first if statement, you have 5 seconds")
time.sleep(10)

for stateCode in dflist2: # loops through dflist and populates conditionals according to code
	
	#pya.typewrite( "{{SubsidiaryPlanId}} = " )

	statement = "AND {{State}} IN " + stateCode

	pyperclip.copy(statement)  #copies code into clipboard
	pya.hotkey('command', 'v') # pastes the code 

	for x in range(4):
		pya.press('tab')

	pya.press('down')
	pya.press('space')



		
