def get_all_textBox(listSlides, data_textBox):
    for slide in listSlides:
        list_shape_textBox = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                textBox_value =  shape.text_frame.text  
                list_shape_textBox.append(textBox_value.replace(" ", ""))
        
        data_textBox.update({slide.slide_id : list_shape_textBox})
    return data_textBox