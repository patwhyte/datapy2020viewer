import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
import os
import random
import time          #remove if we don't use
import sys
import pyfiglet 

############### put try here to verify excel file
os.chdir()
wb = openpyxl.load_workbook('Book1.xlsx', read_only=True)
sheet = wb['Sheet1']

closeApp = False
validInteger = False


def dashboardInfo():
    print("Getting information ...")
    rowLimit = False
    rowCounter = 2
    numCol = 1
    salesCol = 10
    cogsCol = 11
    profitCol = 12

    chanCount = 0
    entCount = 0
    govCount = 0
    midCount = 0
    smallCount = 0

    chanSalesTot = 0
    chanCogsTot = 0
    chanProfitTot = 0
    entSalesTot = 0
    entCogsTot = 0
    entProfitTot = 0
    govSalesTot = 0
    govCogsTot = 0
    govProfitTot = 0
    midSalesTot = 0
    midCogsTot = 0
    midProfitTot = 0
    smallSalesTot = 0
    smallCogsTot = 0
    smallProfitTot = 0
    
    while not rowLimit:
        
        if str(sheet.cell(row = rowCounter, column = numCol).value) == 'Channel Partners':
            chanCount += 1
            chanSalesTot += int(sheet.cell(row = rowCounter, column = salesCol).value)
            chanCogsTot += int(sheet.cell(row = rowCounter, column = cogsCol).value)
            chanProfitTot += int(sheet.cell(row = rowCounter, column = profitCol).value)
            rowCounter += 1

        elif str(sheet.cell(row = rowCounter, column = numCol).value) == 'Enterprise':
            entCount += 1
            entSalesTot += int(sheet.cell(row = rowCounter, column = salesCol).value)
            entCogsTot += int(sheet.cell(row = rowCounter, column = cogsCol).value)
            entProfitTot += int(sheet.cell(row = rowCounter, column = profitCol).value)
            rowCounter += 1

        elif str(sheet.cell(row = rowCounter, column = numCol).value) == 'Government':
            govCount += 1
            govSalesTot += int(sheet.cell(row = rowCounter, column = salesCol).value)
            govCogsTot += int(sheet.cell(row = rowCounter, column = cogsCol).value)
            govProfitTot += int(sheet.cell(row = rowCounter, column = profitCol).value)
            rowCounter += 1

        elif str(sheet.cell(row = rowCounter, column = numCol).value) == 'Midmarket':
            midCount += 1
            midSalesTot += int(sheet.cell(row = rowCounter, column = salesCol).value)
            midCogsTot += int(sheet.cell(row = rowCounter, column = cogsCol).value)
            midProfitTot += int(sheet.cell(row = rowCounter, column = profitCol).value)
            rowCounter += 1
     
        elif str(sheet.cell(row = rowCounter, column = numCol).value) == 'Small Business':
            smallCount += 1
            smallSalesTot += int(sheet.cell(row = rowCounter, column = salesCol).value)
            smallCogsTot += int(sheet.cell(row = rowCounter, column = cogsCol).value)
            smallProfitTot += int(sheet.cell(row = rowCounter, column = profitCol).value)
            rowCounter += 1
      
        else:
            rowLimit = True
            rowCounter -= 2

    salesTotal =  chanSalesTot  + entSalesTot  + govSalesTot  + midSalesTot  + smallCount
    cogsTotal =   chanCogsTot   + entCogsTot   + govCogsTot   + midCogsTot   + smallCogsTot
    profitTotal = chanProfitTot + entProfitTot + govProfitTot + midProfitTot + smallProfitTot
    
    print('****************************************************************************')
    print('There are' + '\u0332'.join(" " + str(rowCounter))+ ' records in the excel file.')
    print('----------------------------------------------------------------------------')
    print()
    print('Segment\t\t\tSales\t\tCost of Sales\t\tProfit')
    print('Channel Partners\t$',str(chanSalesTot).rjust(8," "),'\t$', str(chanCogsTot).rjust(11," "), '\t\t$', str(chanProfitTot).rjust(8," "))
    print('Enterprise\t\t$', str(entSalesTot).rjust(8," "),'\t$', str(entCogsTot).rjust(11," "), '\t\t$', str(entProfitTot).rjust(8," "))
    print('Government\t\t$', str(govSalesTot).rjust(8," "),'\t$', str(govCogsTot).rjust(11," "), '\t\t$', str(govProfitTot).rjust(8," "))
    print('Midmarket\t\t$', str(midSalesTot).rjust(8," "),'\t$', str(midCogsTot).rjust(11," "), '\t\t$', str(midProfitTot).rjust(8," "))
    print('Small Business\t\t$', str(smallSalesTot).rjust(8," "),'\t$', str(smallCogsTot).rjust(11," "), '\t\t$', str(smallProfitTot).rjust(8," "))
    print('============================================================================')
    print('Final Totals\t\t$', str(salesTotal).rjust(8," "),'\t$', str(cogsTotal).rjust(11," "), '\t\t$', str(profitTotal).rjust(8," "))      
    print()
#################################################################

def productListings():
    print("Getting information ...")
    rowLimit = False
    rowCounter = 2
    numCol = 3
    unitsCol = 5
    profitCol = 12

    amaCount = 0
    carCount = 0
    monCount = 0
    pasCount = 0
    velCount = 0
    vttCount = 0

    amaUnits = 0
    carUnits = 0
    monUnits = 0
    pasUnits = 0
    velUnits = 0
    vttUnits = 0

    amaProfit = 0
    carProfit = 0
    monProfit = 0
    pasProfit = 0
    velProfit = 0
    vttProfit = 0
    
    
    while not rowLimit:
        
        if str(sheet.cell(row = rowCounter, column = numCol).value) == 'Amarilla':
            amaCount += 1
            amaUnits += int(sheet.cell(row = rowCounter, column = unitsCol).value)
            amaProfit += int(sheet.cell(row = rowCounter, column = profitCol).value)
            rowCounter += 1


        elif str(sheet.cell(row = rowCounter, column = numCol).value) == 'Carretera':
            carCount += 1
            carUnits += int(sheet.cell(row = rowCounter, column = unitsCol).value)
            carProfit += int(sheet.cell(row = rowCounter, column = profitCol).value)
            rowCounter += 1

        elif str(sheet.cell(row = rowCounter, column = numCol).value) == 'Montana':
            monCount += 1
            monUnits += int(sheet.cell(row = rowCounter, column = unitsCol).value)
            monProfit += int(sheet.cell(row = rowCounter, column = profitCol).value)
            rowCounter += 1

        elif str(sheet.cell(row = rowCounter, column = numCol).value) == 'Paseo':
            pasCount += 1
            pasUnits += int(sheet.cell(row = rowCounter, column = unitsCol).value)
            pasProfit += int(sheet.cell(row = rowCounter, column = profitCol).value)
            rowCounter += 1

        elif str(sheet.cell(row = rowCounter, column = numCol).value) == 'Velo':
            velCount += 1
            velUnits += int(sheet.cell(row = rowCounter, column = unitsCol).value)
            velProfit += int(sheet.cell(row = rowCounter, column = profitCol).value)
            rowCounter += 1
     
        elif str(sheet.cell(row = rowCounter, column = numCol).value) == 'VTT':
            vttCount += 1
            vttUnits += int(sheet.cell(row = rowCounter, column = unitsCol).value)
            vttProfit += int(sheet.cell(row = rowCounter, column = profitCol).value)
            rowCounter += 1
      
        else:
            rowLimit = True
            rowCounter -= 2

    regionTotal = amaCount  + carCount  + monCount  + pasCount  + velCount  + vttCount
    unitTotal   = amaUnits  + carUnits  + monUnits  + pasUnits  + velUnits  + vttUnits
    profitTotal = amaProfit + carProfit + monProfit + pasProfit + velProfit + vttProfit    

    print('****************************************************************************')
    print('There are' + '\u0332'.join(" " + str(rowCounter))+ ' regions reporting.')
    print('-----------------------------------------------------------------------------')
    print()
    print('Product\t\tReporting\tUnits Sold\t\tProfit')
    print('Amarilla\t',str(amaCount).rjust(8," "),'\t', str(amaUnits).rjust(9," "), '\t\t$', str(amaProfit).rjust(8," "))
    print('Carretera\t', str(carCount).rjust(8," "),'\t', str(carUnits).rjust(9," "), '\t\t$', str(carProfit).rjust(8," "))
    print('Montana\t\t', str(monCount).rjust(8," "),'\t', str(monUnits).rjust(9," "), '\t\t$', str(monProfit).rjust(8," "))
    print('Paseo\t\t', str(pasCount).rjust(8," "),'\t', str(pasUnits).rjust(9," "), '\t\t$', str(pasProfit).rjust(8," "))
    print('Velo\t\t', str(velCount).rjust(8," "),'\t', str(velUnits).rjust(9," "), '\t\t$', str(velProfit).rjust(8," "))
    print('VTT\t\t', str(vttCount).rjust(8," "),'\t', str(vttUnits).rjust(9," "), '\t\t$', str(vttProfit).rjust(8," "))
    print('============================================================================')
    print('Final Totals\t', str(regionTotal).rjust(8," "),'\t', str(unitTotal).rjust(9," "), '\t\t$', str(profitTotal).rjust(8," "))      
    print()
    print('****************************************************************************')

#### need excel manual here

#### need db push here
    
##########################################################
def optionDetails():
    print('****************************************************************************')
    print('Option 1 - This will display the total sales and profits of each segement.')
    print('Option 2 - This will display the total units and profit by product.')
    print('Option 3 - Not available...')
    print('Option 4 - Not available...')
    contUser = input("Press any key to continue.")
    print('****************************************************************************')
    print()
    
#############################################################
print('****************************************************************************')
print('****************************************************************************')
appName = pyfiglet.figlet_format("  pyDataViewer", font = "slant"  ) 
print(appName) 
print('****************************************************************************')
print('****************************************************************************')



while not closeApp:
    while not validInteger:
        print('Please choose an option below.')
        print('Option 1: Sales Dashboard - Enter 1')
        print('Option 2: Product Lisitngs - Enter 2')
        print('Option 3: Manual Excel Entry - Enter 3')
        print('Option 4: Export to db - Enter 4')
        print('Option 5: Exit the Program - Enter 99')
    
        userChoice = input('Which task do you want to do?  ')
        if userChoice.isdigit():
            userChoice = int(userChoice)
            
            #checking if the value is a choice in the menu
            if (userChoice > 0 and userChoice < 5) or userChoice == 99 or userChoice ==8:
                validInteger = True
                if userChoice == 1:
                    validInteger = False
                    dashboardInfo()
                    
                elif userChoice == 2:
                    validInteger = False
                    productListings()
                    
                elif userChoice == 3:
                    closeApp = True

                elif userChoice == 4:
                    closeApp = True
                    
                elif userChoice == 8:
                    validInteger = False
                    optionDetails()
                    
                else:
                    closeApp = True
            else:
                print('Please enter a number between 1 - 4 or 99 to EXIT. Press 8 to see option details.')
                    
        else:
            print('Please enter a number between 1 - 4 or 99 to EXIT. Press 8 to see option details.')

print('Exiting app')
wb.close()











