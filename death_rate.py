from xlrd import open_workbook


book = open_workbook('death_rate_final.xlsx')
sheet = book.sheet_by_index(0)
json_output=open("data.json","w")

starting_row_data_index = 3
starting_col_data_index = 0
country_starting_index = 4
country_end_index = sheet.ncols

def death_of_json(row,col,col_country):
    print('inside recurssion')
    if(sheet.cell(row,col).value == ""):
        print('inside null box')
        print(str(row) + '-' + str(col))
        if(sheet.cell(row,col-1).value != ""):
            print(str(row) + '-' + str(col))
            json_output.write('}],')
            death_of_json(row,col-1,col_country)
        else :
            json_output.write('}]}]')
            json_output.write(',')
            print(str(row) + '-' + str(col))
            death_of_json(row,col-2,col_country)
    elif (sheet.cell(row,col+1).value == "" or col == country_starting_index - 1):
        print(str(row) + '-' + str(col))
        if(row == sheet.nrows - 1):
            json_output.write('"' + str(sheet.cell(row,col).value) + '": "' + str(sheet.cell(row,col_country).value) + '"')
            for column in range (0,col):
                json_output.write('}]')
        else:
            if ((sheet.cell(row,col).value == sheet.cell(row + 1,col).value) and row!=129):
                print(str(row) + '-' + str(col))
                json_output.write('"' + str(sheet.cell(row,col).value))
                if(col<3):
                    json_output.write('":[{"count": "' + str(sheet.cell(row,col_country).value) + '",')
                death_of_json(row+1,col+1,col_country)
                
            elif (sheet.cell(row,col).value != sheet.cell(row + 1,col).value):
                print(str(row) + '-' + str(col))
                json_output.write('"' + str(sheet.cell(row,col).value) + '": "' + str(sheet.cell(row,col_country).value) + '"')
                if(sheet.cell(row+1,col).value != ""):
                    json_output.write(',')
                death_of_json(row+1,col,col_country)   


json_output.write('{"deathrate":[')    
for country_index in range(country_starting_index,country_end_index):
    print('Json started')
    json_output.write('{')
    json_output.write('"Country": "' + str(sheet.cell(0,country_index).value) + '",')
    json_output.write('"Code": "' + str(sheet.cell(1,country_index).value) + '",')
    death_of_json(starting_row_data_index,starting_col_data_index,country_index)
    json_output.write('}')
    if(country_index< sheet.ncols-1):
        json_output.write(',')
json_output.write(']}')


