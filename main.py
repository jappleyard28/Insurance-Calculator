import datetime
from datetime import date
import openpyxl


def inflation_factor(from_date, to_date, effective_Dates, percentage_values):
    print("From Date: " + from_date.strftime("%x"))
    print("To Date: " + to_date.strftime("%x"))

    # print(type(Effective_Dates[0].date()))
    date1 = date(2021, 6, 15)
    date2 = date(2021, 6, 20)

    n1 = -100  # array position (i)
    n2 = -101  # array position (j)
    for i in range(len(effective_Dates) - 1):
        if ((effective_Dates[i].date() <= from_date) and (from_date <= effective_Dates[i + 1].date())):
            n1 = i
    for j in range(len(effective_Dates) - 1):
        if ((effective_Dates[j].date() <= to_date) and (to_date <= effective_Dates[j + 1].date())):
            n2 = j

    # initialise inflation factor to 1
    inflation_factor = 1

    print((to_date - from_date).days + 365)
    if (i == j):
        inflation_factor = pow((1 + percentage_values[i]), ((to_date - from_date).days / 365.25))

    else:
        inflation_factor = pow(inflation_factor * (1 + percentage_values[i]),
                               (((effective_Dates[i + 1] - 1).date() - from_date) / 365.25))
        for n in range(i + 1, j - 1):
            inflation_factor = pow((inflation_factor * (1 + percentage_values[n])),
                                   (((effective_Dates[n + 1] - 1).date() - effective_Dates[n].date()) / 365.25))
        inflation_factor = pow((inflation_factor * (1 + percentage_values[j])),
                               ((to_date - effective_Dates[j].date()) / 365.25))

    # TO DO - DOUBLE CHECK IT PASSES ALL ERROR CHECKING
    # checks and errors
    if not (to_date > from_date):
        print("please check there are no claims with dates later than the treaty effective date")
    if not (to_date <= (effective_Dates[-1] + datetime.timedelta(days=365)).date()):
        print(
            "There aren't sufficient values in the corresponding trend factors being used to calculate the inflation factor")
    if not (from_date >= effective_Dates[0].date()):
        print("There are not sufficient historic trend values to produce the inflation factor")

    print("--------------------------------------------------")

    return inflation_factor


def linear_interpolation(X, X1, X2, Y1, Y2):
    Y = Y1 + ((Y1 - Y2) / (X1 - X2)) * (X - X1)
    return Y


def fix_column(column_name):
    # loop through this column and delete blank cells and cells that aren't numerical values
    a = 0
    while (a < len(column_name)):
        if (column_name[a] is None) or not (isinstance(column_name[a], (int, float))):
            del column_name[a]
        else:
            a = a + 1


def sliding_scale(file_name):
    # open excel file
    wb2 = openpyxl.load_workbook(file_name,
                                 data_only=True)  # data_only returns the values instead of formulae for cells
    sheet2 = wb2.active

    gross_revenue = sheet2['J4'].value # gross_revenue input
    if not (isinstance(gross_revenue, (int, float))):
        print("Invalid input")
        return -1
    else:
        column_b = sheet2['B']  # tuple
        column_c = sheet2['C']
        column_d = sheet2['D']
        column_e = sheet2['E']

        # it's printing the starting blank tiles and the title tile
        low_range_list = [cell.value for cell in column_b]
        high_range_list = [cell.value for cell in column_c]
        incr_rate_list = [cell.value for cell in column_d]
        cum_prem_list = [cell.value for cell in column_e]

        # remove all blank rows and rows that aren't numerical values from all the columns I'll be using
        fix_column(low_range_list)
        fix_column(high_range_list)
        fix_column(incr_rate_list)
        fix_column(cum_prem_list)

        # calculate exposure_band
        exposure_band = -1
        for i in range(len(low_range_list)):
            if (gross_revenue < low_range_list[i]):
                exposure_band = i
                break
        if (exposure_band == -1):
            print("Price is out of range, please try inputting a different value")
            return -1
        else:
            # calculate cumulative premium before current band
            cum_prem_below = cum_prem_list[0] # initialise to lowest cumulative premium
            # set cumulative premium below current band to the correct value
            for i in range(len(cum_prem_list)):
                if (i == exposure_band - 2): # -2 because first value is at index 0 and it also selects cum prem from one band below
                    cum_prem_below = cum_prem_list[i]

            # calculate mex exposure below band
            max_exposure_below = high_range_list[0] # initialise to lowest cumulative premium
            # set cumulative premium below current band to the correct value
            for i in range(len(high_range_list)):
                if (i == exposure_band - 2): # -2 because first value is at index 0 and it also selects cum prem from one band below
                    max_exposure_below = high_range_list[i]

            # calculate rate in band
            rate_in_band = incr_rate_list[0]
            for i in range(len(incr_rate_list)):
                if (i == exposure_band - 1):
                    rate_in_band = incr_rate_list[i]

            # calculate base premium
            base_premium = cum_prem_below + (gross_revenue - max_exposure_below) * rate_in_band

            print("exposure band: " + str(exposure_band))
            print("cum prem below band: " + str(cum_prem_below))
            print("max exposure below band: " + str(max_exposure_below))
            print("rate in band: " + format_to_percentage(rate_in_band))
            print("base premium: " + str(base_premium))
            return base_premium

def format_to_percentage(num):
    return str(round(num * 100, 4)) + "%"

if __name__ == '__main__':
    # open inflation factor excel file
    wb = openpyxl.load_workbook("1_Inflation factor (parameters).xlsx")
    sheet = wb.active
    first_column = sheet['A']
    second_column = sheet['B']

    dates = [cell.value for cell in first_column[1:]]  # '1:' makes it ignore the title row
    percentages = [cell.value for cell in second_column[1:]]  # '1:' makes it ignore the title row

    # Function 1
    print("Calculate inflation factor")
    print(inflation_factor(date(2005, 2, 22), date(2002, 7, 28), dates, percentages))

    # Function 2
    print("\nLinear interpolation")
    print("X: 4, Y: " + str(linear_interpolation(4, 2, 6, 4, 7)))

    # Function 3
    print("\nUse sliding scale to calculate cumulative premium")
    print(sliding_scale("3_Sliding Scale Function.xlsx"))
