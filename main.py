import pandas as pd
import os
import sys
import re


def int_converter(value):
    try:
        return int(value)
    except ValueError:
        return value


if __name__ == '__main__':
    GoalXlsx = "Task.xlsx"
    current_directory = os.getcwd()
    GoalPath = os.path.join(current_directory, GoalXlsx)
    try:
        with open(GoalPath):
            print(f"The file '{GoalPath}' loads successfully.")
    except FileNotFoundError:
        print("Error:The file Task.xlsx do not exist.Please rename the excel file to Task.xlsx")
        print("Exception!Press Any to break...")
        input()
        sys.exit()

    df1 = pd.read_excel('Task.xlsx')
    df = pd.read_excel('Task.xlsx', converters={col: int_converter for col in range(df1.shape[1])})
    result_path = 'ResultList'
    os.makedirs(result_path, exist_ok=True)
    # print(df.head())
    TableEnd = df.shape[1]
    # print(f"{TableStart},{TableEnd}")
    # print("请输入attachment从哪个字母列开始（大写，仅支持单个字母）：")
    ColumName = df.columns.tolist()
    TableStart = ord('G') - ord('A')

    # print("请输入数据表的机器号所在列的字母列名（大写，仅支持单个字母）：")
    MCONoPattern = r"VehicleSerialNO"
    ShipLevelPattern = r"Shiplevel"
    AttachmentPattern = r'\b\d{7}\b'
    # MachineNo = ord('E') - ord('A')
    # print("MCONo:", MachineNo)
    # print("请输入数据表的ShipLevel号所在列的字母列名（大写，仅支持单个字母）：")

    # ShipLevel = ord('B') - ord('A')
    # print("ShipLevel:", ShipLevel)
    string_list = [str(item) for item in ColumName]
    for indexC, elem in enumerate(string_list):
        if re.search(MCONoPattern, elem):
            MachineNo = indexC
        if re.search(ShipLevelPattern, elem):
            ShipLevel = indexC
        if re.search(AttachmentPattern,elem):
            TableStart = indexC
            break
    print("MCONo:", MachineNo)
    print("ShipLevel:", ShipLevel)
    print("attachment:", TableStart)
    for indexR, row in df.iterrows():
        row_data = row.tolist()
        # print(row_data)

        if not pd.isna(row_data[MachineNo]):
            FileName = "ResultList\\" + str(row_data[MachineNo]) + str(".txt")
            # print("OK")
            # print(FileName)
            with open(FileName, 'w') as file:
                file.write(str(row_data[ShipLevel]) + "\n")
                # print("OK")
                # print(row_data)
                for i in range(TableStart, TableEnd):
                    if row_data[i] == 1:
                        # print(str(df.columns[i]))
                        file.write(str(df.columns[i])[:7] + "\n")
    print("Success!Press Any to continue...")
    input()
    '''
        if pd.isna(row_data[MachineNo]):
            print("OK")
    '''

    '''
        if not pd.isna(cell_value):
            FileName = "ResultList\\" + str(df.at[indexR, MachineNo]) + str(".txt")
            # print(FileName)
            with open(FileName, 'w') as file:
                for indexR, cell_value in row.iteritems():
                    if cell_value == 1:
                        print(f"Row {indexR}: Cell value is 1")

            print("File 'output.txt' has been created and written.")
    '''
    # print("Row data:",row)
