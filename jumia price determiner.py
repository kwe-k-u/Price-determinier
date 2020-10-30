import pandas as pd


class PriceDeterminer:

    def __init__(self):
        self.path = "C:/Safe/Duala/Afro Kiosk/Documents/Afro Kiosk.xlsx"
        self.sheet = None
        self.sheetname = "Product list"
        self.jumiaRange ={"Selling price" : "Profit"}
        self.regularRange = {"Selling Price" : "Profit"}

    def openfile(self):
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
        file = pd.ExcelFile(self.path)
        self.sheet = file.parse(self.sheetname)

        return self.sheet


    def writePrices(self,index, jumia, regular):
        try:
            if jumia >0:
                self.sheet["I"][index] = jumia
        except:
            print("indeex failed to save")

        try:
            if regular >0:
                self.sheet["G"][index] = regular
        except:
            print(index, "failed to save")





    def jumiaProfit (self, commission, device_cost, delivery, contribution, selling_price):
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


    def regularProfit(self, selling_price, device_cost, delivery):
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


    def getPurchasePrice(self, index):
        column = "F"
        return self.sheet[column][index]



    def getDevice(self, index):
        '''
        Returns the name  of the device at the specified index

        Parameters
        ----------
        index : TYPE
            DESCRIPTION.

        Returns
        -------
        TYPE
            DESCRIPTION.

        '''


        column = "B"
        return self.sheet[column][index]




    def priceRange(self, index):
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
        currentPrice = self.getPurchasePrice(index)
        priceLimit = self.getPurchasePrice(index) * 1.5
        delivery = 28
        contribution = 5
        commission = input("What is the rate for " + self.getDevice(index)+ ": ")


        while currentPrice < priceLimit:
            jprofit = self.jumiaProfit(commission, self.getPurchasePrice(index), delivery, contribution, currentPrice) #jumia profit
            rprofit = self.regularProfit(currentPrice, self.getPurchasePrice(index), delivery) #regular profit


            #Add prices to list if selling price returns a profit
            if jprofit > 0:
                self.jumiaRange[currentPrice] = jprofit
            if rprofit > 0:
                self.regularRange[currentPrice] = rprofit

            currentPrice *= 1.05
            currentPrice = round(currentPrice,2)



    def displayPrices(self, index):
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
        print("Purchase price is ", self.getPurchasePrice(index))
        print("\n\nBelow are the prices for " + self.getDevice(index))

        #saving Jumia price keys in a list
        jkeys =[]
        for k in self.jumiaRange:
            jkeys.append(k)

        #saving Regular price keys in a list
        rkeys = []
        for r in self.regularRange:
            rkeys.append(r)


        #displaying jumia prices
        print("Purchase price is ", self.getPurchasePrice(index))
        print("\nDisplaying the device's suggested prices for Jumia")
        for index in range(len(jkeys)):
            print(index, jkeys[index], self.jumiaRange[jkeys[index]], sep = "\t\t")

        jselIndex = int(input ("Enter the index for your preferred price and profit: "))




        #displaying regular prices
        print("\nDisplaying the device's suggested regular prices")
        for index in range(len(rkeys)):
            print(index, rkeys[index], self.regularRange[rkeys[index]], sep = "\t\t")



        rselIndex = int(input ("Enter the index for your preferred price and profit: "))

        self.writePrices(index, jkeys[jselIndex], rkeys[rselIndex])

        #resetting the price and profit lists
        jkeys =[]
        rkeys = []
        jselIndex = None
        rselIndex = None







    def automatedPriceSelection(self, percentage): #percentage is preferred rate of return
        '''
        Selects a price that returns a profit which is equal or higher than the specified
        preferred profit percentage.
        If none matches criteria, the the price that returns
        The highest profit is selected

        Returns
        -------
        None.

        '''


        self.sheet = self.openfile() #Open spreadsheet

        #loop through all devices
        for index in range(8,90):
            self.priceRange(index)
            newJumiaPrice = 0;
            newRegularPrice = 0;

            upperlimit = self.getPurchasePrice(index) * ((percentage + 5)/100)
            lowerlimit = self.getPurchasePrice(index) * ((percentage - 5)/100)

            #looping through suggested prices for amount that fits criteria (JUMIA)
            for key, profit in self.jumiaRange.items():
                if profit == "Profit":
                    continue
                elif profit >= lowerlimit and profit <= upperlimit:
                    newJumiaPrice = key
                else:
                    try:
                        newJumiaPrice = key
                    except:
                        print("Couldn't save " + self.getDevice(index))





            #looping through suggested prices for amount that fits criteria (REGULAR)
            for key, profit in self.regularRange.items():
                if profit == "Profit":
                    continue
                elif profit >= lowerlimit and profit <= upperlimit:
                    newRegularPrice = key
                else:
                    try:
                        newRegularPrice = key
                    except:
                        print("Couldn't save " + self.getDevice(index))



            #todo remove strings from dictionary to reduce bigO

            #saving to
            self.writePrices(index, newJumiaPrice, newRegularPrice)

            #resetting price ranges
            self.jumiaRange ={"Selling price" : "Profit"}
            self.regularRange = {"Selling Price" : "Profit"}





    def manualPriceSelection(self):

        self.sheet = self.openfile()

        for index in range(8,90):
            self.priceRange(index)
            self.displayPrices(index)
            self.jumiaRange ={"Selling price" : "Profit"}
            self.regularRange = {"Selling Price" : "Profit"}

        self.save()






    def save(self):
        '''

        Saves the editted excel work sheet. NOTE: erases previous formatting

        Returns
        -------
        None.

        '''
        self.sheet.to_excel(self.path,
                 index=False,
                 sheet_name= self.sheetname,)



    def main(self):
        self.automatedPriceSelection(10)
        self.save()




obj = PriceDeterminer()
obj.main()
