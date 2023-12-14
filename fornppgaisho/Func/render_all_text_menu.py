def render_text_menu(listSlides,slide_menu_id):
    object_menu = {}
    slide = listSlides.get(slide_menu_id)
    for shape in slide.shapes:
        if shape.has_table:
            tbl = shape.table
            for row_index in range(2,len(tbl.rows)):
                object_row = {"1.5タ": [],"3.0タ": []}
                index_check_15 = (0,1,2,3,4,8,9)
                index_check_30 = (0,1,5,6,7,8,9)
                for i in index_check_15:
                    object_row["1.5タ"].append(tbl.cell(row_index,i).text_frame.text.replace(" ", ""))
                for j in index_check_30:
                    object_row["3.0タ"].append(tbl.cell(row_index,j).text_frame.text.replace(" ", ""))
                object_menu.update({tbl.cell(row_index,0).text_frame.text : object_row})
    return object_menu


