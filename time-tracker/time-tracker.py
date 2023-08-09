import calendar
import os
import sys

import pandas as pd
from matplotlib import cm

# def highlight_month(s):

# 	# Get array of colors
# 	c_vec = cm.viridis(range(12))

# 	date_dict = {month: index for index, month in enumerate(calendar.month_abbr) if month}

# 	# print(date_dict)
# 	# print(s)
# 	# exit()
# 	color = date_dict[s]

# 	return f'background-color: {color}'


def main():
    # Define column names
    misc_names = ["Day", "Month"]
    project_names = ["AI/ML User Group", "ECP", "LLNL", "Green Computing"]
    pto_names = ["PTO", "Holiday", "Floating Holiday"]

    # start and end dates
    start_date = pd.Timestamp("10/01/2022")
    end_date = pd.Timestamp("9/30/2023")
    d_range = pd.date_range(start=start_date, end=end_date)

    # Create dataframe
    df = pd.DataFrame(
        index=d_range.date,
        columns=misc_names + project_names + pto_names,
    )

    # Initialize the day of the week names
    df["Day"] = d_range.day_name()
    df["Month"] = d_range.month_name()

    # print(df)
    # df.style.apply(highlight_month, subset=["Month"])

    # Save the dataframe as an excel file
    # book_name = os.path.join("sheets", "FY22_timesheet.xlsx")
    writer = pd.ExcelWriter(
        os.path.join("sheets", "FY22_timesheet.xlsx"), engine="openpyxl"
    )
    # with pd.ExcelWriter(book_name, mode="w", engine="openpyxl") as writer:

    for month in d_range.month_name().unique():
        df_tmp = df[df["Month"] == month]
        df_tmp.to_excel(excel_writer=writer, sheet_name=month)

    writer.close()


if __name__ == "__main__":
    main()
