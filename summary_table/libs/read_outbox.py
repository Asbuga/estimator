import pandas as pd


filename = "summary\\data\\outbox\\121_du.xls"

# Work with objects.
objects = pd.read_excel(filename, sheet_name=3)
objects = objects.iloc[:, [2, 5]]
new_objects_columns = ["Номер глави", "Найменування об'єктного кошторису"]
objects.columns = new_objects_columns

# Work with smetas.
smetas = pd.read_excel(filename, sheet_name=4)
smetas = smetas.iloc[:, [2, 4, 5, 6]]
new_smetas_columns = ["Номер глави", "Номер п/п ЛК в об'єкті",
    "Кошторисний номер", "Найменування ЛК"]
smetas.columns = new_smetas_columns

# Work with positions.
positions = pd.read_excel(filename, sheet_name=5)
positions = positions.iloc[:, [7, 7, 1, 2, 4, 5, 7, 8, 9]]
new_positions_columns = ["Розділ", "Примітка", "Номер п/п ЛК в об'єкті",
    "Тип позиції", "Номер позиції в ЛК/в РРР", "Шифр позиції",
    "Наименування", "Одиниці виміру", "Кількість"]
positions.columns = new_positions_columns

# Create columns with divisions and notes.
for index, row in positions.iterrows():

    index_before = index - 1 if index >= 1 else index
    index_2before = index_before - 1 if index_before >= 1 else index_before
    
    check = {
        "index": index - 1 > 0,
        "divisions": row["Тип позиції"].strip() in ["R", "Z", "А", "Б"],
        "note": "Y" in row["Тип позиції"],
        "ls": (positions.at[index, "Номер п/п ЛК в об'єкті"] 
               == positions.at[index_before, "Номер п/п ЛК в об'єкті"]),
        "divisions&note": (positions.at[index_before, "Розділ"] 
               == positions.at[index_2before, "Розділ"])
    }

    if check["divisions"]:
        positions.at[index, "Розділ"] = positions.at[index, "Наименування"]
    
    elif (not check["divisions"]
          and check["index"] 
          and check["ls"]):
        
        positions.at[index, "Розділ"] = positions.at[index_before, "Розділ"]

    else:
        positions.at[index, "Розділ"] = ""
    
    if check["note"]:
        positions.at[index, "Примітка"] = positions.at[index, "Наименування"]
    
    elif (not check["note"] 
          and check["index"] 
          and check["ls"] 
          and check["divisions&note"]
        ):
        
        positions.at[index, "Примітка"] = positions.at[index_before, "Примітка"]

    else:
        positions.at[index, "Примітка"] = ""

# Join dataframe.
smetas = pd.merge(objects, smetas, on="Номер глави")
smetas = pd.merge(smetas, positions, on="Номер п/п ЛК в об'єкті")

# filter empty values.
smetas = smetas.loc[smetas["Шифр позиції"].notna()]
smetas = smetas.iloc[:, [0, 1, 2, 3, 4, 5, 6, 8, 9, 10, 11, 12]]

# Save result.
output_file = "summary_table.xlsx"
smetas.to_excel(output_file, sheet_name="summary", index=False)
