def classify_slide(data):
    slides_part_id= []
    for k,v in data.items():
        #check silde chua table
        if('周波数値の要約表' in v ):
            slide_table_menu_id = k

        for textbox in v:
            if(('以下は細部の応力値です' in textbox) ):
                slides_part_id.append(k)
                break
    return (slide_table_menu_id, slides_part_id)