import pandas as pd

path = "C:/Safe/Duala/Afro Kiosk/Documents/Afro Kiosk.xlsx"
sheet = None
sheetname = "Product list"
jumiaRange ={"Selling price" : "Profit"}
regularRange = {"Selling Price" : "Profit"}

def openfile():
    """

    Returns
    -------
    sheet : TYPE
        DESCRIPTION.

    Open and parse a worksheet to be edited

    Parameters
    ----------
    path : String
        Path to the excel file to be modified.
    sheetname: String
        Name of the worksheet to be worked on

    Returns
    -------
    sheet : TYPE
        Returns a single worksheet with the specified name.

    """
    file = pd.ExcelFile(path)
    sheet = file.parse(sheetname)

    return sheet


def write(column, row, new_value):
    """


    Parameters
    ----------
    column : TYPE
        DESCRIPTION.
    row : TYPE
        DESCRIPTION.
    new_value : TYPE
        DESCRIPTION.

    Returns
    -------
    None.

    Write a new value into a cell

    Parameters
    ----------
    column : String
        The column index of the cell to be edited
    row : Integer
       The row index of the cell to be edited.
    new_value : TYPE
        The value to be inserted in the row.

    Returns
    -------
    None.

    """
    sheet[column][row] = new_value




def jumiaProfit (commission, device_cost, delivery, contribution, selling_price):
    """
    Calculates and returns the expected profit that will be earned when selling with jumia

    Parameters
    ----------
    commission : int
        The percentage that jumia chargers per sale.
    device_cost : float
        The price at which a device is bought.
    delivery : float
        The cost of delivery to the jumia pick up hub.
    contribution : float
        Addition fees charged for vendors.
    selling_price : double
        The price at which the device is to be sold on jumia.

    Returns
    -------
    profit : float
        The profit that will be earned when the device is sold on Juumia at the entered selling price.

    """

    profit = ((100 - float(commission))/100 * float(selling_price)) - device_cost - contribution - delivery
    return round(profit, 2)


def regularProfit(selling_price, device_cost, delivery):
    """
    Calculates and returns the profit to be earned by selling the device at a given price

    Parameters
    ----------
    selling_price : double
        The price at which the device is to be sold.
    device_cost : double
        The cost of the devicee.
    delivery : double
        Cost of delivery of the device.

    Returns
    -------
    profit : double
        The profit to be earned by selling the deivce at the entered price(s).

    """
    profit = selling_price - device_cost - delivery
    return round(profit,2)


def getPurchasePrice(index):
    column = "F"
    return sheet[column][index]



def getDevice(index):
    column = "B"
    return sheet[column][index]




def priceRange(index):
    """
    Generates and checks prices at which a profit will be generated for a device


    Parameters
    ----------
    index : int
        The index of a product in an excel sheet.

    Returns
    -------
    None.

    """
    currentPrice = getPurchasePrice(index)
    priceLimit = getPurchasePrice(index) * 1.5
    delivery = 28
    contribution = 5
    commission = input("What is the rate for " + getDevice(index)+ ": ")


    while currentPrice < priceLimit:
        jprofit = jumiaProfit(commission, getPurchasePrice(index), delivery, contribution, currentPrice) #jumia profit
        rprofit = regularProfit(currentPrice, getPurchasePrice(index), delivery) #regular profit


        #Add prices to list if selling price returns a profit
        if jprofit > 0:
            jumiaRange[currentPrice] = jprofit
        if rprofit > 0:
            regularRange[currentPrice] = rprofit

        currentPrice *= 1.05
        currentPrice = round(currentPrice,2)



def displayPrices(index):
    """
    Displays the suggested prices and lets the user select a price=>profit relationship
    they are comfortable with

    Parameters
    ----------
    index : int
        The index of the product in the excel sheet.

    Returns
    -------
    None.

    """
    print("\n\nBelow are the prices for " + getDevice(index))

    #saving Jumia price keys in a list
    jkeys =[]
    for k in jumiaRange:
        jkeys.append(k)

    #saving Regular price keys in a list
    rkeys = []
    for r in regularRange:
        rkeys.append(r)


    #displaying jumia prices
    print("\nDisplaying the device's suggested prices for Jumia")
    for index in range(len(jkeys)):
        print(index, jkeys[index], jumiaRange[jkeys[index]], sep = "\t\t")

    selIndex = int(input ("Enter the index for your preferred price and profit: "))
    print(jumiaRange[jkeys[selIndex]])



    #displaying jumia prices
    print("\nDisplaying the device's suggested regular prices")
    for index in range(len(rkeys)):
        print(index, rkeys[index], regularRange[rkeys[index]], sep = "\t\t")



    selIndex = int(input ("Enter the index for your preferred price and profit: "))
    print(regularRange[rkeys[selIndex]])

def save():
    """
    Saves the editted excel work sheet. NOTE: erases previous formatting
    """
    sheet.to_excel(path,
             index=False,
             sheet_name=sheetname,)





# =============================================================================
#                   MAIN
# =============================================================================

sheet = openfile()

for index in range(8,81):
    priceRange(index)
    displayPrices(index)

