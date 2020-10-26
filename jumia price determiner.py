import pandas as pd
#selling -> profit

path = "C:/Safe/Duala/Afro Kiosk/Documents/Afro Kiosk.xlsx"

def openfile(path):
    file = pd.ExcelFile(path)
    sheet = file.parse("Product list")
    return sheet


def read(sheet):
    sheet["F"][6]
    #product name (B,8) purchase price (F,8) jumia price

def write(sheet):
    sheet["B"][1] = 2
    print(sheet["B"][1])
    sheet.to_excel(path,
             index=False,
             sheet_name="Product list")

# =============================================================================
#     file = pd.ExcelWriter(path)
#     file.write_cells(["sd"],"Product list", "B", 2)
#     file.save()
# =============================================================================

def jumiaProfit(commission,device_cost, delivery, contributions, selling_price):
    profit = ((100 - commission)/100 *selling_price) - device_cost - delivery - contributions
    return profit

def regularProfit(selling_price, device_cost, delivery):
    profit = selling_price -device_cost - delivery
    return profit

def loopDevices(sheet):
    column = "B"
    numRow = 8

    for row in range(numRow,81):
        print(sheet[column][row])



def loopPurchasePrice(sheet):
    column = "F"
    numRow = 8

    for row in range(numRow,81):
        print(sheet[column][row])



def loopSellingPrice(sheet):
    column = "F"
    numRow = 8

    for row in range(numRow,81):
        print(sheet[column][row])



def loopJumiaPrice(sheet):
    column = "B"
    numRow = 8

    for row in range(numRow,81):
        print(sheet[column][row])
# =============================================================================
#                   MAIN
# =============================================================================
# if the new price is not more than 1.5 times cost price
#todo calculate transaction fees

sheet = openfile(path)
write(sheet)
# =============================================================================
# loopDevices(sheet)
# loopJumiaPrice(sheet)
# loopPurchasePrice(sheet)
# loopSellingPrice(sheet)
# =============================================================================
