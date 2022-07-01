import pandas as pd 
import os
from urllib.request import urlopen
import re
class Data:
    list_data = ['', 0, 0, 0, 0, 0]
    link = ""
    save_Excel_path=""
    save_text_path=''
    sheet_name="sheet1"
    fx=open(save_text_path,'w')
    url_f = urlopen(link)
    myfile = url_f.readlines()

    columns=[  'Date',
                'Sap',
                'Air_Humidity',
                'Air_Temperature',
                'Soil_Humidity',
                'Soil_Temperature']




    df = pd.DataFrame({'Date': [],
                        'Sap': [],
                        'Air_Humidity': [],
                        'Air_Temperature': [],
                        'Soil_Humidity': [],
                        'Soil_Temperature': []})




    i=0
    def sap_temp(DN_f):
        DN=float(DN_f)
        s=127.6-0.006045*DN + 0.000000126*(DN**2) -0.00000000000115*(DN**3)
        return s



    list_len = len(myfile)
    nexttime=False
    nbre_of_lignes=0
    for line in myfile:
        nbre_of_lignes+=1
        current_idx = myfile.index(line)
        new_line=line.decode().strip()
        line_split=re.split(',|;',new_line)
        #print(line_split)
        list_end = list_len - current_idx
        if ((line_split[1] == "") and (line_split[3] == "4B")):
            if(nexttime ==True):
                #print(list_data, "= ", i)
                fx.write(str(list_data))
                fx.write("\n")
                c0=list_data[0]
                c1=list_data[1]
                c2=list_data[2]
                c3=list_data[3]
                c4=list_data[4]
                c5=list_data[5]
                df.loc[len(df.index)] = [c0,c1,c2,c3,c4,c5]


            list_data = ['', None, None, None, None, None]
            # list_data = ['', 0, 0, 0, 0, 0]
            i += 1
            list_data[0] = line_split[0]

            ref_1 = True
            ref_2 = True
        if ((line_split[1] == "62190380") and (line_split[3] == "55") and (ref_1 == True)):
            t_ref_0 = sap_temp(line_split[5])
            t_heat_0 = sap_temp(line_split[6])
            t_ref_1 = sap_temp(line_split[18])
            t_heat_1 = sap_temp(line_split[19])
            air_hum = float(line_split[10])
            air_temp = float(line_split[11]) / 10
            Ton = t_heat_1 - t_ref_1
            Toff = t_heat_0 - t_ref_0
            sap = 12.95 * ((Ton / (Ton - Toff)) - 1) * 27.777
            list_data[1] = sap
            list_data[2] = air_hum
            list_data[3] = air_temp

            ref_1 = False
        if ((line_split[1] == "64210015") and (line_split[3] == "55") and (ref_2 == True)):
            soil_hum = float(line_split[10])
            soil_temp = float(line_split[11]) / 10
            list_data[4] = soil_hum
            list_data[5] = soil_temp
            ref_2 = False
        if list_end != 1:
            nexttime=True

        elif list_end == 1:
            nexttime=False
            print("number of row =  ",i)
            fx.write(str(list_data))
            fx.write("\n")
            c0 = list_data[0]
            c1 = list_data[1]
            c2 = list_data[2]
            c3 = list_data[3]
            c4 = list_data[4]
            c5 = list_data[5]
            df.loc[len(df.index)] = [c0, c1, c2, c3, c4, c5]



    with pd.ExcelWriter(save_Excel_path) as writer:
        df2 = df.style.highlight_null(null_color='red')

        df2.to_excel(writer, sheet_name=sheet_name, index=False)
        for col in columns:
            col_idx = df.columns.get_loc(col)
            writer.sheets[sheet_name].set_column(col_idx, col_idx, len(col))
            writer.save()



print("upload sucessful")

