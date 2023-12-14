from Func.Find_slide_menu import classify_slide
from Func.change_color_text_tbl import change_color
from Func.getAll_textbox import get_all_textBox
from Func.render_all_text_menu import render_text_menu

from pptx import Presentation


link_pp = 'C://Users//ADMIN//Desktop//code//python//Tkinter//fornppgaisho//DummyPPGaisho//DummyData.pptx'
path_save = 'C://Users//ADMIN//Desktop//code//python//Tkinter//fornppgaisho//DummyPPGaisho//result.pptx'
    
psr = Presentation(link_pp)
listSlides = psr.slides

#get all data textbox
data_textBox = get_all_textBox(listSlides, {})
# print(data_textBox) 

#classify slides
(slide_menu_id, slides_part_id) = classify_slide(data_textBox)
print("slide_menu_id: ", slide_menu_id) 
print("slides_part_id: ", slides_part_id) 


# với loại baeng chia 1.5j và 3.0j
    #get table trong  slide menu
obj_menu = render_text_menu(listSlides,slide_menu_id)
print(obj_menu)

    #work in slide part
part_slide = listSlides.get(257)
        #get table in part slide
for shape in part_slide.shapes:
        if shape.has_table:
            tbl = shape.table
            rows = tbl.rows
            check_index = (0,1,3,5,6,4,7)
            for i in range(1,len(rows)):
                list_cells = []
                for j in check_index:
                    cell = rows[i].cells[j]
                    if(cell.is_spanned == True):
                        list_cells.append(rows[i-1].cells[j].text_frame.text.replace(" ", ""))
                    if(cell.is_spanned == False):
                        list_cells.append(cell.text_frame.text.replace(" ", ""))             
                        
                list_cells[0] =list_cells[0].split("-")[0]
                #check xem dang la 1.5j hay 3.0j
                
                
                current_obj = obj_menu[list_cells[0]]
                
                def compare_menu(obj, kind):
                    list_text_table = obj[kind]
                    if(list_text_table == list_cells ):
                        pass
                    if(list_text_table != list_cells ):
                        list_part = list_cells.copy()
                        print("list_part: ", list_part)
                        print("list_text_table: ", list_text_table)
                        
                        
                        # for part_cell in list_part:
                        #     for menu_cell in list_text_table:
                        #         if(menu_cell == part_cell ):
                        #             list_part.remove(part_cell)
                        # print("list_part: ", list_part)
                        
                                    
                    # list_cells
                    
                if("1.5" in rows[i].cells[2].text_frame.text):
                    compare_menu(current_obj, "1.5タ")
                if("3.0" in rows[i].cells[2].text_frame.text):
                    print("check object 3.0")
                    
                # print("list_cells: ", list_cells)


psr.save(path_save)