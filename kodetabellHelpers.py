# -*- coding: utf-8 -*-
"""
Created on Thu Jul  4 13:55:16 2019

@author: Håkon Weydahl (weyhak@banenor.no)

Inneholder klasser som kan ta en mengde kodetabeller og hente ut informasjonen. 
Bibliotektet som snakker med excel (xlrd) er testet kun på .xls-filer. 
Dersom kodetabellen er i det nyere .xlsx-formatet kan kodetabellen lagres på
nytt i gammelt format. 

Objekter:
    -   Baliseoversikt(): "Permen" med alle kodetabellene du er interessert i. 
        Innholder en liste over alle kodetabellene
    -   Kodetabell(): Hvert enkelt regneark, inneholder en liste over alle 
        balisegruppene på arket
    -   Balisegruppe(): Den enkelte bgruppe, inneholder en liste over alle 
        balisene i gruppa
    -   Balise(): En enkelt balise
    
"""

from pathlib import Path
import re
from numpy import gradient
import pandas as pd
import xlrd
import xlsxwriter

###########
# Klasser #
###########

class Baliseoversikt:
    def __init__(self):
        self.alle_ktab = []
        
    def ny_mappe(self, folder_path):
        for file in self.__getXLSfileList(folder_path):
            self.alle_ktab.append(Kodetabell(file))

    def ny_fil(self, filename):
        if Path(filename).suffix == ".xls":
            self.alle_ktab.append(Kodetabell(filename))
        else:
            print(f"Error: Feil filformat på '{filename}'. Må være .xls")

    def finn_materiell(self, save_path):
        # kontroller path
        save_path = Path(save_path)
        save_path.parent.mkdir(parents=True, exist_ok=True)

        # les data fra Balise() objekter
        data = {}
        for ktab in self.alle_ktab:
            for bgrp in ktab.balise_group_obj_list:

                # nye lister
                balise_plugs = []
                balises = {"F": 0, "Y": 0, "Z" : 0, "Y/Z" : 0}
                coder_cards = {i : 0 for i in range(1,16)}

                # for hvert balise
                for balise in bgrp.baliser:
                    for reg in [balise.x_reg, balise.y_reg, balise.z_reg]:
                        if len(reg) == 1:
                            balise_plugs += reg
                        else:
                            coder_cards += reg
                            print(f"Koder kort?", reg)

                    if (len(balise.x_reg) == 1 and len(balise.y_reg) == 1 and len(balise.z_reg) == 1):
                        balises["F"] += 1
                    elif (len(balise.y_reg) == 1 and len(balise.z_reg) > 1):
                        balises["Y"] += 1
                    elif (len(balise.y_reg) > 1 and len(balise.z_reg) == 1):
                        balises["Z"] += 1
                    elif (len(balise.y_reg) > 1 and len(balise.z_reg) > 1):
                        balises["Y/Z"] += 1
                    else:
                        print(f"Error: {balise.id1}{balise.id1} ukjent balise type?")

                # finn type
                bgrp_type = bgrp_speed = None
                if (bgrp.sign_type == "ERH") and (bgrp.tilstander[0]["vent"] != None):
                    bgrp_type = "ERH"
                    bgrp_speed = bgrp.tilstander[0]["vent"]
                elif (bgrp.sign_type == "EH") and (bgrp.tilstander[0]["kjor"] != None) and (bgrp.tilstander[0]["motr_type"] == "SEH"):
                    bgrp_type = "EH/SEH"
                elif (bgrp.sign_type == "EH") and (bgrp.tilstander[0]["kjor"] != None) and (bgrp.tilstander[0]["motr_type"] != "SEH"):
                    bgrp_type = "EH"
                elif (bgrp.sign_type == "SEH"):
                    bgrp_type = "SEH"
                else:
                    print(f"Error: Finne ikke skilt for {bgrp.id1}{bgrp.id2}.")

                            
                # data[bgrp.id1 + bgrp.id2] = [balise_plugs]
                data[bgrp.id1 + bgrp.id2] = {
                    "raw_static" : balise_plugs,
                    "raw_dynamic" : coder_cards,
                    "balises" : balises,
                    "type" : bgrp_type,
                    "speed" : bgrp_speed,
                }

        # transformer til materiell tabell
        for bgrpID in data:
            plug_count = [0 for i in range(16)]
            for i, item in enumerate(plug_count):
                for e in data[bgrpID]["raw_static"]:
                    if (e == i):
                        plug_count[i] += 1
            # data[bgrpID].append(plug_count)
            data[bgrpID]["sorted_static"] = plug_count

        # skriv til excelark
        # cols = 1 + 4 + 15
        start_col = 0
        start_row = 1


        workbook = xlsxwriter.Workbook(save_path)
        worksheet_C = workbook.add_worksheet("Materiell Maximo")
        worksheet_A = workbook.add_worksheet("Materiell per objekt")
        worksheet_B = workbook.add_worksheet("Materiell total")

        # A-tabell
        table_data_A = []
        for bgrpID in data:
            table_row = [bgrpID]
            table_row.append(data[bgrpID]["type"])
            table_row.append(data[bgrpID]["speed"])
            table_row += data[bgrpID]["balises"].values()
            table_row += data[bgrpID]["sorted_static"]
            table_data_A.append(table_row)

        # kolonne overskrifter baliser
        column_data = [
            {"header" : "ID"},
            {"header" : "Type"},
            {"header" : "Skilt"}
            ]
        # column_data.append(
        #     {"header" : "ID"},
        #     {"header" : "Type"},
        #     {"header" : "Hastighet"},
        #     )
        for element in balises.keys():
            column_data.append({
                "header" : f"{element}",
                "total_function" : "sum"
                })

        # kolonne overskrifter kodeplugger
        for i in range(16):
            column_data.append({
                "header" : f"# {i}",
                "total_function" : "sum"
                })

        # lag tabellen
        end_col = start_col + len(column_data) - 1
        worksheet_A.add_table(start_row, start_col, start_row+len(table_data_A)+2, end_col, {
            "data" : table_data_A,
            "columns" : column_data,
            "total_row" : True,
            })

        worksheet_A.set_column(start_col+1, end_col, 6)

        # lag merged header format
        header_format = workbook.add_format({
            "bold" : 1,
            "font_color" : "white",
            "align" : "center",
            "valign" : "vcenter",
            "fg_color" : "#4F81BD",
            "border" : 2,
            "border_color" : "white"
        })

        # lag merged header baliser
        worksheet_A.merge_range(
            start_row-1, start_col+3, start_row-1, start_col+3+3,
            "Baliser",
            header_format
            )

        # lag merged header kodeplugg
        worksheet_A.merge_range(
            start_row-1, start_col+3+4, start_row-1, end_col,
            "Kodeplugger",
            header_format
            )
        
        # kolonnebredde 
        worksheet_A.set_column(0, 2, 10)
        worksheet_A.set_column(3, end_col, 6)



        # lag totaler for B-tabell
        totals = {}
        for key in balises.keys():
            totals[f"{key} baliser"] = 0
        for i in range(16):
            totals[f"Kodeplugg {i}"] = 0

        # for hver balisegruppe
        for bgrpID in data:

            # summer baliser
            for key in data[bgrpID]["balises"]:
                totals[f"{key} baliser"] += data[bgrpID]["balises"][key]

            # summer kodeplugger
            for i in range(16):
                totals[f"Kodeplugg {i}"] += data[bgrpID]["sorted_static"][i]
        
        # lag B-tabell
        table_data_B = []
        for key in totals:
            table_data_B.append([key, totals[key]])
        worksheet_B.add_table(
            0, 0, len(table_data_B), 1,
            {
                "data" : table_data_B,
                "columns" : [
                    {"header" : "Materiell"},
                    {"header" : "Antall"},
                ],
            }
        )

        # sett bredere A-kolonne
        worksheet_B.set_column(0, 0, 20)


        """ 
        C-tabell 
        
        """

        bData = [
            ['Balise type F u/kabel - NY', 105990, 0, "F"],
            ['Balise type F u/kabel - BRUKT', 104662, 0, ""],
            ['Balise Y Styrbar ADtranz - NY', None, 0, "Y"],
            ['Balise Y Styrbar ADtranz - BRUKT', None, 0, ""],
            ['Balise Z Styrbar ADtranz - NY', 105026, 0, "Z"],
            ['Balise Z Styrbar ADtranz - BRUKT', 103986, 0, ""],
            ['Balise Y/Z Styrbar ADtranz - NY', 106231, 0, "Y/Z"],
            ['Balise Y/Z Styrbar ADtranz - BRUKT', 105940, 0, ""],
        ]

        pData = [
            ['Balisepropper Kodeord 0', 107604,0],
            ['Balisepropper Kodeord 0 BRUKT', 107190,0],
            ['Balisepropper Kodeord 1', 103987,0],
            ['Balisepropper Kodeord 1 BRUKT', 104550,0],
            ['Balisepropper Kodeord 2', 108177,0],
            ['Balisepropper Kodeord 2 BRUKT', 106700,0],
            ['Balisepropper Kodeord 3', 108178,0],
            ['Balisepropper Kodeord 3 BRUKT', 106701,0],
            ['Balisepropper Kodeord 4', 105027,0],
            ['Balisepropper Kodeord 4 BRUKT', 103988,0],
            ['Balisepropper Kodeord 5', 104551,0],
            ['Balisepropper Kodeord 5 BRUKT', 105114,0],
            ['Balisepropper Kodeord 6', 105076,0],
            ['Balisepropper Kodeord 6 BRUKT', 106423,0],
            ['Balisepropper Kodeord 7', 105163,0],
            ['Balisepropper Kodeord 7 BRUKT', None,0],
            ['Balisepropper Kodeord 8', 103989,0],
            ['Balisepropper Kodeord 8 BRUKT', None,0],
            ['Balisepropper Kodeord 9', 105991,0],
            ['Balisepropper Kodeord 9 BRUKT', 105164,0],
            ['Balisepropper Kodeord 10', 103990,0],
            ['Balisepropper Kodeord 10 BRUKT', 104552,0],
            ['Balisepropper Kodeord 11', 107191,0],
            ['Balisepropper Kodeord 11 BRUKT', 107478,0],
            ['Balisepropper Kodeord 12', 104553,0],
            ['Balisepropper Kodeord 12 BRUKT', 105115,0],
            ['Balisepropper Kodeord 13', 107605,0],
            ['Balisepropper Kodeord 13 BRUKT', 107192,0],
            ['Balisepropper Kodeord 14', 107479,0],
            ['Balisepropper Kodeord 14 BRUKT', None,0],
            ['Balisepropper Kodeord 15', 106702,0],
            ['Balisepropper Kodeord 15 BRUKT', None,0],
            ]

        sData = {
            30 : ['Skilt, Tallplate 2 og 3', 105259,0],
            50 : ['Skilt, Tallplate 4 og 5', 104957,0],
            70 : ['Skilt, Tallplate 6 og ', 105159,0],
            90 : ['Skilt, Tallplate 8 og 9', 104704,0],
            110 : ['Skilt, Tallplate 10 og 11', 106167,0],
            130 : ['Skilt, Tallplate 12 og 13', 104705,0],
        }

        mData = {
            "stativ" : ['Skilt, Stativ for Signal 69A og 69B', 106580,0],
            "feste" : ['Balisefeste for midlertidig hastighetsreduksjon', 104083,0],
            "mmerke" : ['Skilt, Signal 68D, Markeringsmerke', 107076,0],
            "mhast" : ['Skilt, Signal 69A og 69B, Midlertidig hastighet', 107424,0],
        }
        
        # lag totaler for C-tabeller
        # for hver balisegruppe
        for bgrpID in data:

            # summer linjemateriell
            # hastighetsskilt
            if data[bgrpID]["speed"] is not None:
                for key in sData:
                    # print(type(data[bgrpID]["speed"]), data[bgrpID]["speed"], "\t", type(key), key)
                    if data[bgrpID]["speed"] <= key:
                        sData[key][2] += 1
                        break
            
            # skilt-materiell
            if data[bgrpID]["type"] == "ERH":
                mData["stativ"][2] += 1
                mData["mmerke"][2] += 0
                mData["mhast"][2] += 1
            elif data[bgrpID]["type"] == "EH/SEH":
                mData["stativ"][2] += 1
                mData["mmerke"][2] += 1
                mData["mhast"][2] += 1
            elif data[bgrpID]["type"] == "EH":
                mData["stativ"][2] += 1
                mData["mmerke"][2] += 1
                mData["mhast"][2] += 0
            elif data[bgrpID]["type"] == "SEH":
                mData["stativ"][2] += 1
                mData["mmerke"][2] += 0
                mData["mhast"][2] += 1
            else:
                print(f"Error: Ingen type for '{bgrpID}'")


            # summer balise-type
            for bName in bData:
                for key in data[bgrpID]["balises"]:
                    if bName[3] == key:
                        bName[2] += data[bgrpID]["balises"][key]

            # summer kodeplugger
            for i in range (0, len(pData), 2):
                pData[i][2] += data[bgrpID]["sorted_static"][i//2]
        
        # antall stativer for C2-tabell
        for bName in bData:
            mData["feste"][2] += bName[2]

        # lag C1-tabell
        table_data_C1 = []
        for key in sData:
            table_data_C1.append([sData[key][0], sData[key][1], sData[key][2]])
        for key in mData:
            table_data_C1.append([mData[key][0], mData[key][1], mData[key][2]])
        
        worksheet_C.add_table(
            0, 0, len(table_data_C1), 3,
            {
                "data" : table_data_C1,
                "columns" : [
                    {"header" : "Materiell Linje"},
                    {"header" : "Artikkelnummer"},
                    {"header" : "Antall"},
                    {"header" : "Kommentar"},
                ],
            }
        )

        # lag C2-tabell
        table_data_C2 = []
        for bName in bData:
            table_data_C2.append([bName[0], bName[1], bName[2]])
        table_data_C2.append([None, None, None])
        for pName in pData:
            table_data_C2.append([pName[0], pName[1], pName[2]])
        
        worksheet_C.add_table(
            len(table_data_C1)+2, 0, len(table_data_C1)+2+len(table_data_C2), 3,
            {
                "data" : table_data_C2,
                "columns" : [
                    {"header" : "Materiell Signal"},
                    {"header" : "Artikkelnummer"},
                    {"header" : "Antall"},
                    {"header" : "Kommentar"},
                ],
            }
        )

        # sett bredere kolonner
        worksheet_C.set_column(0, 0, 40)
        worksheet_C.set_column(1, 1, 15)
        worksheet_C.set_column(3, 3, 30)


        # rydd opp
        workbook.close()

            
    # hvordan oversikten printes
    def __str__(self):
        balisegrupper_df = PD_table(self.alle_ktab)
        print(balisegrupper_df.balise_df)
        return ""

    # Finner alle .xls filer i angitt mappe
    def __getXLSfileList(self, folder_path):
        # import os
        # xls_files = []
        # (_, _, filenames) = next(os.walk(folder_path))
    
        # for file in filenames:
        #     if file.lower().endswith(".xls"):
        #         xls_files.append(folder_path + "\\" + file)

        xls_files = list(Path(folder_path).rglob("*.xls"))

        # sjekker om filer er funnet, slutter hvis ikke
        if len(xls_files) == 0:
            print(f"Ingen .XLS-filer funnet i '{folder_path}'")
            exit()
        else:
            print("Antall .XLS-filer funnet: {}" .format(len(xls_files)))
            return xls_files
    

class Kodetabell:
    def __init__(self, filepath):
        self.filepath = filepath
        self.balise_group_obj_list = [] # liste over alle balisegrupper på arket
        
        # Hvilke kolonner i excel-arket som definerer en tilstand
        # <navn> : <excel-kolonne>
        self.ktab_cols = {
                "H" : "F",
                "F/H" : "G",
                "F" : "H",
                "kjor" : "I",
                "vent" : "J",
                "p-avstand" : "K",
                "b-avstand" : "L",
                "fall" : "M",
                "PX" : "AP", "PY" : "AQ", "PZ" : "AR", # p-balise
                "AX" : "AS", "AY" : "AV", "AZ" : "AX", # a-balise
                "BX" : "AZ", "BY" : "BA", "BZ" : "BC", # b-balise
                "CX" : "BE", "CY" : "BF", "CZ" : "BG", # c-balise
                "NX" : "BH", "NY" : "BI", "NZ" : "BJ", # n-balise
                "motr_type" : "BP",
                "motr_hast" : "BQ"
                }
             
        # initiering starter her
        self.__les_kodetabell()
    
    def __les_kodetabell(self):
        print(self.filepath)
        self.wbook = xlrd.open_workbook(self.filepath) # åpner excel workbook
        self.wb_sheet = self.wbook.sheet_by_index(0) # aktiverer sheet nr 0
        
        self.__definer_balisegrupper() # lager balise_group_obj_list
        
        for bgruppe in self.balise_group_obj_list:
            bgruppe = self.__definer_tilstander(bgruppe)
            bgruppe.kodere = self.__tell_kodere(bgruppe)
            # print(bgruppe.id2, "\n", bgruppe.tilstander) # kun for debugging. printer output
    
    
    # søker etter balisegrupper i kodetabellen
    def __definer_balisegrupper(self):
        for group_row in range(5,42):            
            # Lager balise objekt med __init__ info 
            if (self.wb_sheet.cell(group_row,1).ctype==0) or \
            (self.wb_sheet.cell(group_row,2).ctype==0 and
             self.wb_sheet.cell(group_row,3).ctype==0):# and
            #  self.wb_sheet.cell(group_row,4).ctype==0):
                continue
            else:
                self.balise_group_obj_list.append(Balisegruppe(
                    self.wb_sheet.cell_value(group_row,1), # sign_type
                    self.wb_sheet.cell_value(group_row,2).strip(), # id1
                    self.wb_sheet.cell_value(group_row,3).strip(), # id2
                    self.__clean_KM(self.wb_sheet.cell_value(group_row,4)), # km
                    self.wb_sheet.cell_value(5,0), # ktab retning
                    self.wb_sheet.cell_value(50,90), # s_nr
                    group_row, # første rad nr
                    self.__last_row(group_row) # siste rad nr
                ))

    # finner alle definerte tilstander for en balisegruppe
    def __definer_tilstander(self, group_obj):        
        # search_col, returnerer en liste for hver kolonne        
        kolonne_dict = {}
        for key in self.ktab_cols:
            value = self.__search_col(
                            col_name(self.ktab_cols[key]),
                            group_obj
                            )
            if value != None:
                kolonne_dict.update({key : value})
                
        # lager en linje per tilstand
        tilstand_list = []
        row_span = group_obj.last_row - group_obj.first_row + 1
        for row in range(row_span):            
            tilstand_linje = {}
            for key in kolonne_dict:
                # print(row, key, kolonne_dict[key][row]) # debugging
                tilstand_linje[key] = kolonne_dict[key][row]
            togvei_celle = self.wb_sheet.cell_value(
                    group_obj.first_row + row,
                    col_name("CB")
                    )
            
            # kopier over eventuelt innhold fra celle med togvei info
            if togvei_celle != "":                
                tilstand_linje["togvei"] = self.wb_sheet.cell_value(
                        group_obj.first_row + row,
                        col_name("CB")
                        )
            
            tilstand_list.append(tilstand_linje)
        group_obj.tilstander = tilstand_list
        
        # lager Balise objekt med info om koding
        for litra in ["P", "A", "B", "C"]:
            if litra + "X" in kolonne_dict:
                group_obj.baliser.append(Balise(
                        litra,
                        kolonne_dict[litra + "X"],
                        kolonne_dict[litra + "Y"],
                        kolonne_dict[litra + "Z"]
                        ))
                
        # sette km på balisene        
        if ("A" in group_obj.retning):
            retning = 1
        else:
            retning = -1
            
        offset = 8 # hvor mange meter fra hsign til første balise

        # relativ distanse mellom baliser i gruppen
        dist = {
            "P" : -3,
            "A" : 0,
            "B" : 3, 
            "C" : 6,
        }
        
        for balise in group_obj.baliser:
            if group_obj.type == "H.sign":
            #     egen_gruppe = [balise.rang for balise in group_obj.baliser]
            #     for i, bokstav in enumerate(egen_gruppe):
            #         if bokstav == balise.rang:
            #             balise.km = group_obj.km + (3 * i - offset) * retning
                balise.km = group_obj.km + (dist[balise.rang] - offset) * retning                    
            else:
                # egen_gruppe = [balise.rang for balise in group_obj.baliser]
                # print(egen_gruppe)
                # for i, bokstav in enumerate(egen_gruppe):
                #     if bokstav == balise.rang:
                #         print(i)
                balise.km = group_obj.km + dist[balise.rang] * retning                    
                # if balise.rang is "P":
                #     balise.km = group_obj.km - 3 * retning
                # else:
                #     for i, bokstav in enumerate(["A", "B", "C"]):
                #         if bokstav == balise.rang:
                #             balise.km = group_obj.km + 3 * i * retning
        # def slutt
        return group_obj
    
    # leter i kommentarfeltet etter gyldige koderbenevninger, returnerer liste
    def __tell_kodere(self, group_obj):
        
        # gyldige navn på kodere:
        koder_benevning = (
        "FSK[1-9]*"
        "|HSK[1-9]*"
        "|DSK[1-9]*"
        "|VK[ZY1-9]*"
        "|PK[ZY1-9]*"
        "|BK[ZY1-9]*"
        "|CK[ZY1-9]*"
        "|REP\.*K[1-9]*"
        "|RSK[1-9]*"
        )
        
        koder_list = []        
        for row in range(group_obj.first_row, group_obj.last_row + 1):
            kommentar_celle = self.wb_sheet.cell_value(row, col_name("CA"))
            if kommentar_celle == "":
                continue
            else:
                match_obj = re.findall(koder_benevning, kommentar_celle, re.I | re.X)
                if match_obj:
                    [koder_list.append(item) for item in match_obj]
        return koder_list
            
            
        
    # leser en kolonne fra top til bunn og kopierer innhold
    # returner liste dersom normal, returner None dersom kolonna er tom
    def __search_col(self, col, group_obj):
        
        row_code = []
        row = group_obj.first_row # første rad i siste balise-objekt fra liste
        if (self.wb_sheet.cell(row, col).ctype == 2) or (self.wb_sheet.cell_value(row, col) != ""): # hvis har innhold
            row_code.append(
                    self.wb_sheet.cell_value(row, col) # les kode fra celle
                    )
        else: # hvis ikke innhold
            # row_code.append(None) # returner liste med None per rad
            return None # returner None i stedet for en liste
        
        if group_obj.first_row == group_obj.last_row:
            return self.__make_int(row_code)
        else:
            for row in range(group_obj.first_row + 1, group_obj.last_row + 1):
                if (self.wb_sheet.cell(row, col).ctype == 2) or (self.wb_sheet.cell_value(row, col) != ""): # hvis har innhold
                    row_code.append(
                            self.wb_sheet.cell_value(row, col) # les kode fra celle
                            )
                else: # hvis ikke innhold
                    row_code.append(row_code[-1]) # kopierer kode fra forrige linje
            return self.__make_int(row_code)
              
    # finner antall rader en balisegruppe strekker seg over
    def __last_row(self, first_row):
        last_row = first_row      
        for key in self.ktab_cols:
            col = col_name(self.ktab_cols[key])
            row = first_row
            while True:
                if (self.wb_sheet.cell(row + 1, col).ctype == 2) or (self.wb_sheet.cell_value(row + 1, col) == "-"): # hvis cellen ikke er tom
                    row += 1
                else:
                    break
            if row > last_row:
                last_row = row
        return last_row

    # del av search_col()
    def __make_int (self, aList):
        newList = []
        for string in aList:
            try:
                newList.append(int(string))
            except:
                newList.append(string)
        if len(aList) != len(newList):
            print("__make_int error")
        return newList
    
    # fjerner rusk fra KM og returnerer en int
    def __clean_KM(self, KM_str):
        from re import findall
        KM_str = str(KM_str)
        if KM_str.isdigit():
            return KM_str
        else:
            try:
                KM_str = "".join(findall("[0-9]", KM_str))
                return int(KM_str)
            except:
                print(KM_str)
                print(findall("[0-9]", KM_str))
                return -1.0


class Balisegruppe:
    def __init__(self, sign_type, id1, id2, km, ktab_retning, s_nr, first_row, last_row):
        self.sign_type = sign_type
        self.id1 = id1
        self.id2 = id2
        self.km = km
        self.ktab_retning = ktab_retning
        self.s_nr = s_nr
        self.first_row = first_row
        self.last_row = last_row
        self.tilstander = None
        self.kodere = []
        self.sim_segment = None # segment dersom den skal brukes i ATC sim
        self.baliser = []
        
        self.finn_retning()
        self.finn_type()
        
        # setter retning avhengig av om id2 er odde er partall
    def finn_retning(self):
        m = re.match("\d+", self.id2[::-1])
        try:
            nr = int(m.group(0)[::-1])
            if nr % 2 == 0:
                self.retning = "B"
            else:
                self.retning = "A"
        except:
            self.retning = "?"

    # klassifiserer etter type balisegruppe        
    def finn_type(self):
        # https://trv.banenor.no/wiki/Signal/Prosjektering/ATC#Baliseidentitet
        tabell_12 = {
                "H.sign": ["_", "M", "O", "S", "Y", "Æ", "Å", "L", "N", "P", "T", "X", "Ø"],
                "D.sign": ["m", "o", "s", "y", "æ", "å", "l", "n", "p", "t", "x", "ø"],
                "F.sign": ["F"],
                "FF": ["Z"],
                # "Rep.": ["R", "U", "V", "W"],
                "Rep.": ["R", "U", "W"], # V er for SVG
                "L": ["L"],
                "SVG/RVG": ["V", "v"],
                "SH": ["S"],
                "H/H(K1)/H(K2)": ["H"],
                "ERH/EH/SEH": ["E"],
                "GMD/GMO/HG/BU/SU": ["G"]
                }
        for key in tabell_12:
            if self.id2[0] in tabell_12[key] or self.id2[1] in tabell_12[key]:
                self.type = key
        
    
    def __str__(self):
        self_str = "{}\t{} {}\t{}\t" .format(self.sign_type, self.id1, self.id2, self.km)
        return self_str


class Balise:
    def __init__(self, rang, x_reg, y_reg, z_reg):
        self.rang = rang # P, A, B, C eller N-balise
        self.x_reg = x_reg # X-ord
        self.y_reg = y_reg
        self.z_reg = z_reg
        self.km = 0
        
    def __str__(self):
        return ("{0}X: {1}\t{0}Y: {2}\t{0}Z: {3}" .format(
                self.rang, 
                self.x_reg, 
                self.y_reg, 
                self.z_reg
                ))    


class PD_table:
    import pandas as pd
    def __init__(self, ktab_liste):
        self.ktab_liste = ktab_liste
        
        self.pd_import = []
        for ktab in self.ktab_liste:
            for balise in ktab.balise_group_obj_list:
                self.pd_import.append({
                    "Sign./type" : balise.sign_type,
                    "Sted" : balise.id1,
                    "ID" : balise.id2,
                    "KM" : balise.km,
                    "Retning" : balise.retning,
                    "Tegning" : balise.s_nr,
                    "Rad nr." : "{}-{}" .format(balise.first_row+1, balise.last_row+1),
                    "Kodere" : len(balise.kodere)
                })
        
        self.balise_df = pd.DataFrame(self.pd_import)
        self.balise_df = self.balise_df[['Retning', 'Sign./type', 'Sted', 'ID', 'KM', 'Tegning', 'Rad nr.', 'Kodere']]
        
    def lagre_excel(self):
        self.balise_df.to_excel("gruppeliste.xlsx")
        
    def print_df(self):
        print(self.balise_df)



##############
# Funksjoner #
##############

# vasker kodeord for å presenteres i excel    
def rens_kodeord(kodeliste):
    
    kodeliste = set(kodeliste)
    
    if "-" in kodeliste:
        kodeliste.remove("-")
        if len(kodeliste) == 0:
            return 1 # kode "1" dersom koding er vilkårlig
    
    if len(kodeliste) == 1:
        return kodeliste.pop()
    else: 
        return ', '.join(map(str, kodeliste))

# Lager excelark med baliser
def skrivBaliseliste(ktabList, wbName):
    import xlsxwriter

    # Lager liste med dictionaries
    baliseDictList = []
    for ktab in ktabList.alle_ktab:
        for bgruppe in ktab.balise_group_obj_list:
            for balise in bgruppe.baliser:                
                baliseDictList.append({
                        "Retning": bgruppe.retning,
                        "Sign./Type": bgruppe.sign_type,
                        "Type": bgruppe.type,
                        "ID_sted": bgruppe.id1, 
                        "ID_type": bgruppe.id2, 
                        "KM_prosjektert": balise.km,
                        "KM_simulering": 0,
                        # "KM_simulering": "=" + lagReferanse(len(baliseDictList)+2, 6-1), # rad og kolonne det skal refereres til
                        # "Segment": evaluerSegment(balise, len(baliseDictList)+2, 7), # for å gjøre P, B, C til referanse
                        "Segment": bgruppe.sim_segment,                        
                        "Rang": balise.rang, 
                        "X-ord": rens_kodeord(balise.x_reg), 
                        "Y-ord": rens_kodeord(balise.y_reg), 
                        "Z-ord": rens_kodeord(balise.z_reg),
                        "Tegning": bgruppe.s_nr, 
                        "Rad nr.": bgruppe.first_row + 1
                        })

    # Lage workbook-objekt
    workbook  = xlsxwriter.Workbook(wbName)
    baliseWorksheet = workbook.add_worksheet("Balisegrupper")

    # Estetikk
    # listContent = workbook.add_format({"align": "center"})
    # tableHeader = workbook.add_format({"bold": True, "border": True, "align": "center"})

    # Definer tabell
    data = makeListOfLists(baliseDictList)
    baliseWorksheet.add_table(0,0, len(data), len(data[0])-1, {
        "data": data,
        "columns": makeHeaders(baliseDictList)
        # "header_row": True
        })
    
    # Opprydding
    workbook.close()
    return

# tar en bokstavkode, gir plass i alfabetet
def alphabet_number(some_char):
    return ord(some_char.upper())-64

def col_name(letter_string):
    sum = 0
    for idx, c in enumerate(reversed(letter_string)):
        sum += 26**idx*alphabet_number(c)
    return sum - 1

def makeListOfLists(DictList):
    return [list(dictionary.values()) for dictionary in DictList]

def makeHeaders(DictList):
    return [{"header": "{}" .format(key)} for key in DictList[0]]

def angiSporsegment(ktabList, trackSegment):
    # iterate through all balise groups
    for ktab in ktabList.alle_ktab:
        for bgroup in ktab.balise_group_obj_list:
            
            # iterate through all IDs for all balise groups
            for key in trackSegment:
                for id in trackSegment[key]:
                    # set balise group sim segment if group is in segment list
                    if (id == bgroup.id1 + bgroup.id2):
                        bgroup.sim_segment = key

# def find_plug(code_list):
#     if len(code_list) == 0:
#         print(f"Error: Ingenting i '{code_list}'.")
#         return None
#     elif len(code_list) == 1:
#         return code_list[0]
#     else:
#         print(f"Error: Flere ting i '{code_list}'.")
#         return None
