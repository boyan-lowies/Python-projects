
import openpyxl as op

wb = op.load_workbook('C:\\Users\\lowie\\Desktop\\pytest.xlsx', data_only=True)
sheetTP = wb['TP current']
sheetD = wb['Data']
sheetN = wb['Next']

statics = "77256, 19790"



#------------------------#
#      Subroutines       #
#------------------------#
 

def search_value_in_column(ws, search_string, column):
    for row in range(1, ws.max_row + 1):
        coordinate = "{}{}".format(column, row)
        if ws[coordinate].value == search_string:
            return row
    return

#------------------------#
#      Drill list        #
#------------------------#

sub_items = ["AF","AH","AJ","AL","AN"]
sub_count = ["AG","AI","AK","AM"]


def drill_list(id, row, buy):
    item = []
    idrow = []
    count = []
    for x in range(0, len(row)):
#        print(x,row[x])
        adress = sub_items[0] + str(row[x])
        z=0   
        while sheetD[adress].value != None:
            y = sheetD[adress].value
            if y not in item:               
                item.append(y) 
                count.append(sheetD[sub_count[z] + str(row[x])].value)
            else:
                index = item.index(y)
                count[index] = count[index] + sheetD[sub_count[z] + str(row[x])].value
            z += 1
            adress = sub_items[z] + str(row[x])

    for x in item:
        if x in buy:
            index = item.index(x)
            del item[index]
            del count[index]
        else:
            idrow.append(search_value_in_column(sheetD,str(x),"AC"))
          

    return item, count, idrow


    


#------------------------#
#         Main           #
#------------------------#



items=[]
itemid=[]
itemc=[]
itemc_row=[]

N=2
adress= "D2"  


while sheetTP[adress].value is not None:
    if sheetTP["A" + str(N)].value == None:
        items.append(sheetTP["D" + str(N)].value)
        x=len(items)-1
        itemid.append(sheetD["A" + str(search_value_in_column(sheetD,items[x],"B"))].value)
        itemc_row.append(search_value_in_column(sheetD,str(itemid[x]),"AC"))
        itemc.append(sheetD["AB" + str(itemc_row[x])].value)
    N += 1
    adress= "D" + str(N)  


buy_id = []
N=2
adress = "O2"

while sheetN[adress].value is not None:
    
    buy_id.append(sheetN[adress].value)
    N += 1
    adress = "O" + str(N)

buy_id.append(statics)

print(len(itemid), len(itemc_row))

list_1 = drill_list(itemid,itemc_row, buy_id)
print(list_1[2], len(list_1[2]))
list_2 = drill_list(list_1[0], list_1[2], buy_id)