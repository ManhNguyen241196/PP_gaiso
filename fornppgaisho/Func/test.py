list_part=  ['#1', 'Chitiet1', '24.5', '244.0', 'X', '300.0', 'AL']
list_text_table=  ['#1', 'Chitiet1', '24.5', '244.5', 'X', '300.0', 'AL']
copy_tbl = list_part.copy()
def hello():
    for index , part_cell in enumerate(list_part):
        for menu_cell in list_text_table:
           if(part_cell == menu_cell):
                copy_tbl.remove(menu_cell)
           if(part_cell != menu_cell):
               pass
            
                
hello()
print(copy_tbl)
