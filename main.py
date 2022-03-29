import sys
import csv
import glob
import os
import pandas as pd
import numpy as np
import time

overhead = .4

while True:
    try:
        #years = [input("Select Year: ")]
        #months = [input("Select Month: ")]
        years = ["2022"]
        months = ["March"]
        #years = ["2017", "2018", "2019", "2020", "2021", "2022"]
        #months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        # get data file names
        x = 0
        y = 0
        y2017 = []
        y2018 = []
        y2019 = []
        y2020 = []
        y2021 = []
        y2022 = []
        y2023 = []
        y2024 = []
        y2025 = []
        y2026 = []
        y2027 = []

        yearList = [y2018, y2019, y2020, y2021, y2022, y2023, y2024, y2025, y2026, y2027]

        excel_files = []

        rentalsList = []
        truckingList = []
        servicesList = []
        laborList = []
        revenueList = []
        equipmentList = []
        commissionList = []

        for i in years:
            x = 0
            for i in months:
                path = f'/Users/tylermorton/SyncDrive/Harvest PM Shared (Jason Holmes)/Event Expense Sheets/{years[y]}/{months[x]}'
                files = os.listdir(path)
                files_xls = [f for f in files if f[-4:] == 'xlsx']
                z = 0
                for f in files_xls:
                    sheet = pd.read_excel(f'/Users/tylermorton/SyncDrive/Harvest PM Shared (Jason Holmes)/Event Expense Sheets/{years[y]}/{months[x]}/{files_xls[z]}')
                    rows = sheet.axes[0]
                    col = sheet.columns
                    subrent = sheet["Unnamed: 2"][7]
                    trucking = sheet["Unnamed: 5"][25]
                    services = sheet["Unnamed: 16"][25]
                    labor = sheet["Unnamed: 6"][7]
                    revenue = sheet["Unnamed: 2"][25]
                    equipment = sheet["Unnamed: 16"][12]
                    commission = sheet["Unnamed: 16"][7]

                    if commission == "Bartle":
                        commissionAmount = int(sheet["Unnamed: 2"][25]) * .15
                        commissionList.append(commissionAmount)
                    else:
                        commissionAmount = 0
                        commissionList.append(commissionAmount)

                    #print(commissionList)
                    if isinstance(subrent, float) or isinstance(subrent, int):
                        if np.isnan(subrent):
                            rentalsList.append(0)
                        else:
                            rentalsList.append(subrent)

                    if isinstance(trucking, float) or isinstance(trucking, int):
                        if np.isnan(trucking):
                            truckingList.append(0)
                        else:
                            truckingList.append(trucking)

                    if isinstance(services, float) or isinstance(services, int):
                        if np.isnan(services):
                            servicesList.append(0)
                        else:
                            servicesList.append(services)

                    if isinstance(labor, float) or isinstance(labor, int):
                        if np.isnan(labor):
                            laborList.append(0)
                        else:
                            laborList.append(labor)

                    if isinstance(revenue, float) or isinstance(revenue, int):
                        if np.isnan(revenue):
                            revenueList.append(0)
                        else:
                            revenueList.append(revenue)

                    if isinstance(equipment, float) or isinstance(equipment, int):
                        if np.isnan(equipment):
                            equipmentList.append(0)
                        else:
                            equipmentList.append(equipment)
                    z += 1
                x += 1
            y += 1



        # print(sum(rentalsList))
        # print(sum(truckingList))
        # print(sum(servicesList))
        # print(sum(laborList))
        # print(sum(revenueList))
    except:
        pass

    try:
        rentalTot = sum(rentalsList)
        truckingTot = sum(truckingList)
        servicesTot = sum(servicesList)
        laborTot = sum(laborList)
        revenueTot = sum(revenueList)
        equipPurchaseTot = sum(equipmentList)
        commissionTot = sum(commissionList)


        laborPer = laborTot / sum(revenueList)
        truckingper = truckingTot / sum(revenueList)
        servicesper = servicesTot / sum(revenueList)
        rentalsper = rentalTot / sum(revenueList)
        equipmentper = equipPurchaseTot / sum(revenueList)
        commissionper = commissionTot / sum(revenueList)

        #print(rentalsper)
        print("\n_____________________________")
        print("\nRevenue Total: $" + str(revenueTot))
        if rentalsper > 0:
            print("\nRental Total: $" + str(round(rentalTot, 2)))
            print("Rentals: " + str(rentalsper)[2:4] + "." + str(rentalsper)[5] + "%")
        else:
            print("\nRental Total: None")
            print("Rentals: None")
        if laborPer > 0:
            print("\nLabor Total: $" + str(round(laborTot, 2)))
            print("Labor: " + str(laborPer)[2:4] + "." + str(laborPer)[5] + "%")
        else:
            print("\nLabor Total: None")
            print("Labor: None")
        if truckingper > 0:
            print("\nTrucking Total: $" + str(round(truckingTot, 2)))
            print("Trucking: " + str(truckingper)[2:4] + "." + str(truckingper)[5] + "%")
        else:
            print("\nTrucking Total: None")
            print("Trucking: None")
        if servicesper > 0:
            print("\nServices Total: $" + str(round(servicesTot, 2)))
            print("Services: " + str(servicesper)[2:4] + "." + str(servicesper)[5] + "%")
        else:
            print("\nServices Total: None")
            print("Services: None")
        if commissionper > 0:
            print("\nCommission Total: $" + str(commissionTot))
            print("Commission: " + str(commissionper)[2:4] + "." + str(commissionper) [5] + "%")
        else:
            print("\nCommission Total: None")
            print("Commission: None")

        print("\nEquipment Purchase Allocation: $" + str(equipPurchaseTot))
        print("Equipment Purchase Allocation: " + str(equipPurchaseTot / revenueTot)[2:5] + "%")
        print("\nProfit: $" + str((revenueTot) - (commissionTot + rentalTot + truckingTot + laborTot + equipPurchaseTot + (revenueTot * overhead))))
        profit = ((revenueTot) - (commissionTot + rentalTot + truckingTot + laborTot + equipPurchaseTot + (revenueTot * overhead)))
        print(profit/revenueTot)
        print("Profit: " + str(profit / revenueTot)[2:4] + "." + str(profit / revenueTot)[4:5] + "%\n")
    except:
        print("No budget sheets for this month!")
        pass
    time.sleep(2)

